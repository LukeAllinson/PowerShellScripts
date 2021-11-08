#Requires -Version 5 -Modules ExchangeOnlineManagement

<#
	.SYNOPSIS
		Name: Export-EXORoomCalendarInformation.ps1
		This gathers Room mailbox calendar processing information and exports to a csv file.

	.DESCRIPTION
		This script connects to EXO and then outputs Room mailbox calendar processing inforamtion to a CSV file.

	.NOTES
		Version: 0.3
        Updated: 08-11-2021 v0.3    Updated filename ordering
        Updated: 19-10-2021 v0.2    Refactored using current script standards
		Updated: <unknown>	v0.1	Initial draft

		Authors: Luke Allinson (github:LukeAllinson)
                 Robin Dadswell (github:RobinDadswell)

    .PARAMETER OutputPath
        Full path to the folder where the output will be saved.
        Can be used without the parameter name in the first position only.

    .PARAMETER MailboxFilter
        Provide a filter to reduce the size of the Get-EXOMailbox query; this must follow oPath syntax standards.
        For example:
        'EmailAddresses -like "*Conference*"'
        'DisplayName -like "*Ground*"'
        'CustomAttribute1 -eq "InScope"'

    .EXAMPLE
        .\Export-EXORoomCalendarInformation.ps1 C:\Scripts\
        Exports Calendar processing information for all Room mailboxes

    .EXAMPLE
        .\Export-EXORoomCalendarInformation.ps1 C:\Scripts\ -MailboxFilter 'Department -eq "R&D"'
        Exports Calendar processing information for Room mailboxes in the the R&D department
#>

[CmdletBinding()]
param
(
    [Parameter(
        Mandatory,
        Position = 0
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateScript(
        {
            if (!(Test-Path -Path $_))
            {
                throw "The folder $_ does not exist"
            }
            else
            {
                return $true
            }
        }
    )]
    [IO.DirectoryInfo]
    $OutputPath,
    [Parameter()]
    [Alias('Filter')]
    [string]
    $MailboxFilter
)

function Get-RoomMailboxInformation ($mailbox)
{
    # Get Mailbox Statistics
    try
    {
        Write-Verbose "Getting Room mailbox information for $($mailbox.PrimarySmtpAddress)"
        $calendarInfo = Get-CalendarProcessing -Identity $mailbox.Guid -WarningAction SilentlyContinue -ErrorAction Stop
    }
    catch
    {
        throw
    }

    # Store everything in an Arraylist
    $roomMailboxInfo = [ordered]@{
        UserPrincipalName     = $mailbox.UserPrincipalName
        DisplayName           = $mailbox.Displayname
        PrimarySmtpAddress    = $mailbox.PrimarySmtpAddress
        AutomateProcessing    = $calendarInfo.AutomateProcessing
        AddOrganizerToSubject = $calendarInfo.AddOrganizerToSubject
        DeleteComments        = $calendarInfo.DeleteComments
        DeleteSubject         = $calendarInfo.DeleteSubject
        RemovePrivateProperty = $calendarInfo.RemovePrivateProperty
        AddAdditionalResponse = $calendarInfo.AddAdditionalResponse
        AdditionalResponse    = $calendarInfo.AdditionalResponse
    }

    Write-Verbose "Completed gathering Room mailbox information for $($mailbox.PrimarySmtpAddress)"
    return [PSCustomObject]$roomMailboxInfo
} #End Function Get-RoomMailboxInformation

### Main Script
# Check if there is an active Exchange Online PowerShell session and connect if not
$PSSessions = Get-PSSession | Select-Object -Property State, Name
if ((@($PSSessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -eq 0)
{
    Write-Verbose 'Not connected to Exchange Online, prompting to connect'
    Connect-ExchangeOnline
}

# Define constants for use later
$timeStamp = Get-Date -Format ddMMyyyy-HHmm
Write-Verbose 'Getting Tenant Name for file name from Exchange Online'
$tenantName = (Get-OrganizationConfig).Name.Split('.')[0]
$outputFile = $OutputPath.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar + 'EXORoomCalendarInformation_' + $tenantName + '_' + $timeStamp + '.csv'
$output = New-Object System.Collections.Generic.List[System.Object]

Write-Verbose "Checking if $outputFile already exists"
if (Test-Path $outputFile -ErrorAction SilentlyContinue)
{
    throw "The file $outputFile already exists, please delete the file and try again."
}

# Define a hashtable for splatting into Get-EXOMailbox
$commandHashTable = @{
    RecipientTypeDetails = 'RoomMailbox'
    ResultSize           = 'Unlimited'
    ErrorAction          = 'Stop'
}

if ($MailboxFilter)
{
    $commandHashTable['Filter'] = $MailboxFilter
}

# Get mailboxes using the parameters defined from the hashtable and throw an error if encountered
try
{
    Write-Verbose 'Getting Room Mailboxes from Exchange Online'
    $mailboxes = @(Get-EXOMailbox @commandHashTable)
}
catch
{
    throw
}
Write-Verbose "There are $mailboxCount Room mailboxes"
$mailboxCount = $mailboxes.Count

if ($mailboxCount -eq 0)
{
    return 'There are no Room mailboxes found using the supplied filters'
}

#  Loop through the list of mailboxes and output the results to the CSV
Write-Verbose 'Beginning loop through all Room mailboxes'
foreach ($mailbox in $mailboxes)
{
    Write-Progress -Id 1 -Activity 'EXO Room Mailbox Calendar Information Report' -Status "Processing $($i) of $($mailboxCount) Mailboxes --- $($mailbox.UserPrincipalName)" -PercentComplete (($i * 100) / $mailboxCount)
    try
    {
        $mailboxInfo = Get-RoomMailboxInformation $mailbox
        $output.Add([PSCustomObject]$mailboxInfo) | Out-Null
    }
    catch
    {
        continue
    }
    finally
    {
        $i++
    }
}

Write-Progress -Activity 'EXO Room Mailbox Calendar Information Report' -Id 1 -Completed
$output | Export-Csv $outputFile -NoClobber -NoTypeInformation -Encoding UTF8

return "Room mailbox information has been exported to $outputfile"
