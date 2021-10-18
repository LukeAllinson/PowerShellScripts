#Requires -Version 5 -Modules ExchangeOnlineManagement

<#
	.SYNOPSIS
		Name: Get-EXOMailboxSizeReport.ps1
		This gathers mailbox size information including primary and archive size and item count.

	.DESCRIPTION
		This script connects to EXO and then outputs Mailbox statistics to a CSV file.

	.NOTES
		Version: 0.3
        Updated: 15-10-2021 v0.3    Refactored for new parameters, error handling and verbose messaging
        Updated: 14-10-2021 v0.2    Rewritten to improve speed, remove superflous information
		Updated: <unknown>	v0.1	Initial draft

		Authors: Luke Allinson, Robin Dadswell

    .PARAMETER OutputPath
        Full path to the folder where the output will be saved.
        Can be used without the parameter name in the first position only.

    .PARAMETER IncludeCustomAttributes
        Include custom and extension attributes; these are not included by default.

    .PARAMETER RecipientTypeDetails
        Provide one or more RecipientTypeDetails values to return only mailboxes of those types in the results. Seperate multiple values by commas.
        Valid values are: DiscoveryMailbox, EquipmentMailbox, GroupMailbox, RoomMailbox, SchedulingMailbox, SharedMailbox, TeamMailbox, UserMailbox.

    .PARAMETER MailboxFilter
        Provide a filter to reduce the size of the Get-EXOMailbox query; this must follow oPath syntax standards.
        For example:
        'EmailAddresses -like "*bruce*"'
        'DisplayName -like "*wayne*"'
        'CustomAttribute1 -eq "InScope"'

    .EXAMPLE
        .\Export-EXOMailboxsizes.ps1 C:\Scripts\
        Exports size information for all mailbox types

    .EXAMPLE
        .\Export-EXOMailboxsizes.ps1 -RecipientTypeDetails RoomMailbox,EquipmentMailbox -Output C:\Scripts\
        Exports size information only for Room and Equipment mailboxes

    .EXAMPLE
        .\Export-EXOMailboxsizes.ps1 C:\Scripts\ -MailboxFilter 'Department -eq "R&D"'
        Exports size information for all mailboxes from the R&D department
#>

[CmdletBinding()]
param
(
    [Parameter(
        Mandatory,
        Position=0
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateScript(
        {
            if (!(Test-Path -Path $_)) {
                throw "The folder $_ does not exist"
            } else {
                return $true
            }
        }
    )]
    [IO.DirectoryInfo]
    $OutputPath,
    [Parameter()]
    [ValidateSet(
        "DiscoveryMailbox",
        "EquipmentMailbox",
        "GroupMailbox",
        "RoomMailbox",
        "SchedulingMailbox",
        "SharedMailbox",
        "TeamMailbox",
        "UserMailbox"
    )]
    [string[]]
    $RecipientTypeDetails,
    [Parameter()]
    [Alias("Filter")]
    [string]
    $MailboxFilter
)

function Get-MailboxInformation ($mailbox) {
    # Get Mailbox Statistics
    try {
        Write-Verbose "Getting mailbox statistics for $($mailbox.PrimarySmtpAddress)"
        $primaryStats = Get-EXOMailboxStatistics -Identity $mailbox.Guid -Properties LastLogonTime -WarningAction SilentlyContinue -ErrorAction Stop
    }
    catch {
        throw
    }

    $primaryTotalItemSizeMB = $primaryStats | Select-Object @{name=”TotalItemSizeMB”; expression={[math]::Round(($_.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1MB),2)}}

    # If an Archive exists, then get Statistics
    if ($mailbox.ArchiveStatus -ne "None") {
        Write-Verbose "Getting archive mailbox statistics for $($mailbox.PrimarySmtpAddress)"
        $archiveStats = Get-EXOMailboxStatistics -Identity $mailbox.Guid -Properties LastLogonTime -Archive -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
        $archiveTotalItemSizeMB = $archiveStats | Select-Object @{name=”TotalItemSizeMB”; expression={[math]::Round(($_.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1MB),2)}}
    }
    # Store everything in an Arraylist
    $mailboxInfo = [ordered]@{
        UserPrincipalName = $mailbox.UserPrincipalName
        DisplayName = $mailbox.Displayname
        PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
        Alias = $mailbox.Alias
        RecipientTypeDetails = $mailbox.RecipientTypeDetails
        LitigationHoldEnabled = $mailbox.LitigationHoldEnabled
        RetentionHoldEnabled = $mailbox.RetentionHoldEnabled
        InPlaceHolds = $mailbox.InPlaceHolds
        ArchiveStatus = $mailbox.ArchiveStatus
    }

    if ($primaryStats) {
        $mailboxInfo["TotalItemSize(MB)"] = $primaryTotalItemSizeMB.TotalItemSizeMB
        $mailboxInfo["ItemCount"] = $primaryStats.ItemCount
        $mailboxInfo["DeletedItemCount"] = $primaryStats.DeletedItemCount
        $mailboxInfo["LastLogonTime"] = $primaryStats.LastLogonTime
    }

    if ($archiveStats) {
        $mailboxInfo["Archive_TotalItemSize(MB)"] = $archiveTotalItemSizeMB.TotalItemSizeMB
        $mailboxInfo["Archive_ItemCount"] = $archiveStats.ItemCount
        $mailboxInfo["Archive_DeletedItemCount"] = $archiveStats.DeletedItemCount
        $mailboxInfo["Archive_LastLogonTime"] = $archiveStats.LastLogonTime
    }

    Write-Verbose "Completed gathering mailbox statistics for $($mailbox.PrimarySmtpAddress)"
    return [PSCustomObject]$mailboxInfo
} #End Function Get-MailboxInformation

### Main Script
# Check if there is an active Exchange Online PowerShell session and connect if not
$PSSessions = Get-PSSession | Select-Object -Property State, Name
if ((@($PSSessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -eq 0) {
    Write-Verbose "Not connected to Exchange Online, prompting to connect"
    Connect-ExchangeOnline
}

# Define constants for use later
$i = 1
$timeStamp = Get-Date -Format ddMMyyyy-HHmm
Write-Verbose "Getting Tenant Name for file name from Exchange Online"
$tenantName = (Get-OrganizationConfig).Name.Split(".")[0]
$outputFile = $OutputPath.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar + $timeStamp + '-' + $tenantName + '-' + 'EXOMailboxSizes.csv'

Write-Verbose "Checking if $outputFile already exists"
if (Test-Path $outputFile -ErrorAction SilentlyContinue) {
    throw "The file $outputFile already exists, please delete the file and try again."
}

$output = [System.Collections.ArrayList]@()

# Define a hashtable for splatting into Get-EXOMailbox
$commandHashTable = @{
    Properties = "LitigationHoldEnabled","RetentionHoldEnabled","InPlaceHolds","ArchiveStatus"
    ResultSize = "Unlimited"
    ErrorAction = "Stop"
}

if ($RecipientTypeDetails) {
    $commandHashTable["RecipientTypeDetails"] = $RecipientTypeDetails
}

if ($MailboxFilter) {
    $commandHashTable["Filter"] = $MailboxFilter
}

# Get mailboxes using the parameters defined from the hashtable and throw an error if no results are returned
try {
    Write-Verbose "Getting Mailboxes from Exchange Online"
    $mailboxes = @(Get-EXOMailbox @commandHashTable)
}
catch {
    throw
}

$mailboxCount = $mailboxes.Count
Write-Verbose "There are $mailboxCount number of mailboxes"

if ($mailboxCount -eq 0) {
    throw "There are no mailboxes found using the filters requested."
}

#  Loop through the list of mailboxes and output the results to the CSV
Write-Verbose "Beginning loop through all mailboxes"
foreach ($mailbox in $mailboxes) {
    Write-Progress -Id 1 -Activity "EXO Mailbox Size Report" -Status "Processing $($i) of $($mailboxCount) Mailboxes --- $($mailbox.UserPrincipalName)" -PercentComplete (($i*100)/$mailboxCount)
    try {
        $mailboxInfo = Get-MailboxInformation $mailbox
        $output.Add([PSCustomObject]$mailboxInfo) | Out-Null
    }
    catch {
        continue
    }
    finally {
        $i++
    }
}

Write-Progress -Activity "EXO Mailbox Size Report" -Id 1 -Completed
$output | Export-Csv $outputFile -NoClobber -NoTypeInformation -Encoding UTF8

return "Mailbox size data has been exported to $outputfile"
