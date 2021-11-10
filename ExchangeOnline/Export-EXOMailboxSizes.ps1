#Requires -Version 5 -Modules ExchangeOnlineManagement

<#
    .SYNOPSIS
        Name: Export-EXOMailboxSizes.ps1
        This gathers mailbox size information including primary and archive size and item count and exports to a csv file.

    .DESCRIPTION
        This script connects to EXO and then outputs Mailbox statistics to a CSV file.

    .NOTES
        Version: 0.8
        Updated: 10-11-2021 v0.8    Added parameter sets to prevent use of mutually exclusive parameters
                                    Disabled write-progress if the verbose parameter is used
        Updated: 10-11-2021 v0.7    Updated to include inactive mailboxes and improved error handling
        Updated: 08-11-2021 v0.6    Fixed an issue where archive stats are not included in output if the first mailbox does not have an archive
                                    Updated filename ordering
        Updated: 19-10-2021 v0.5    Updated to use Generic List instead of ArrayList
        Updated: 18-10-2021 v0.4    Updated formatting
        Updated: 15-10-2021 v0.3    Refactored for new parameters, error handling and verbose messaging
        Updated: 14-10-2021 v0.2    Rewritten to improve speed, remove superflous information
        Updated: <unknown>  v0.1    Initial draft

        Authors: Luke Allinson (github:LukeAllinson)
                 Robin Dadswell (github:RobinDadswell)

    .PARAMETER OutputPath
        Full path to the folder where the output will be saved.
        Can be used without the parameter name in the first position only.

    .PARAMETER InactiveMailboxOnly
        Only gathers information about inactive mailboxes (active mailboxes are not included in results).

    .PARAMETER IncludeInactiveMailboxes
        Include inactive mailboxes in results; these are not included by default.
    
    .PARAMETER RecipientTypeDetails
        Provide one or more RecipientTypeDetails values to return only mailboxes of those types in the results. Seperate multiple values by commas.
        Valid values are: DiscoveryMailbox, EquipmentMailbox, GroupMailbox, RoomMailbox, SchedulingMailbox, SharedMailbox, TeamMailbox, UserMailbox.

    .PARAMETER MailboxFilter
        Provide a filter to reduce the size of the Get-EXOMailbox query; this must follow oPath syntax standards.
        For example:
        'EmailAddresses -like "*bruce*"'
        'DisplayName -like "*wayne*"'
        'CustomAttribute1 -eq "InScope"'

    .PARAMETER Filter
        Alias of MailboxFilter parameter.

    .EXAMPLE
        .\Export-EXOMailboxSizes.ps1 C:\Scripts\
        Exports size information for all mailbox types

    .EXAMPLE
        .\Export-EXOMailboxSizes.ps1 -RecipientTypeDetails RoomMailbox,EquipmentMailbox -OutputPath C:\Scripts\
        Exports size information only for Room and Equipment mailboxes

    .EXAMPLE
        .\Export-EXOMailboxSizes.ps1 C:\Scripts\ -MailboxFilter 'Department -eq "R&D"'
        Exports size information for all mailboxes from the R&D department
#>

[CmdletBinding(DefaultParameterSetName = 'DefaultParameters')]
param
(
    [Parameter(
        Mandatory,
        Position = 0,
        ParameterSetName = 'DefaultParameters'
    )]
    [Parameter(
        Mandatory,
        Position = 0,
        ParameterSetName = 'InactiveOnly'
    )]
    [Parameter(
        Mandatory,
        Position = 0,
        ParameterSetName = 'IncludeInactive'
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
    [Parameter(
        ParameterSetName = 'InactiveOnly'
    )]
    [switch]
    $InactiveMailboxOnly,
    [Parameter(
        ParameterSetName = 'IncludeInactive'
    )]
    [switch]
    $IncludeInactiveMailboxes,
    [Parameter(
        ParameterSetName = 'DefaultParameters'
    )]
    [Parameter(
        ParameterSetName = 'InactiveOnly'
    )]
    [Parameter(
        ParameterSetName = 'IncludeInactive'
    )]
    [ValidateSet(
        'DiscoveryMailbox',
        'EquipmentMailbox',
        'GroupMailbox',
        'RoomMailbox',
        'SchedulingMailbox',
        'SharedMailbox',
        'TeamMailbox',
        'UserMailbox'
    )]
    [string[]]
    $RecipientTypeDetails,
    [Parameter(
        ParameterSetName = 'DefaultParameters'
    )]
    [Parameter(
        ParameterSetName = 'InactiveOnly'
    )]
    [Parameter(
        ParameterSetName = 'IncludeInactive'
    )]
    [Alias('Filter')]
    [string]
    $MailboxFilter
)

function Get-MailboxInformation ($mailbox)
{
    # Get Mailbox Statistics
    Write-Verbose "Getting mailbox statistics for $($mailbox.PrimarySmtpAddress)"
    try
    {
        $primaryStats = Get-EXOMailboxStatistics -Identity $mailbox.Guid -IncludeSoftDeletedRecipients -Properties LastLogonTime -WarningAction SilentlyContinue -ErrorAction Stop
        $primaryTotalItemSizeMB = $primaryStats | Select-Object @{name = 'TotalItemSizeMB'; expression = { [math]::Round(($_.TotalItemSize.ToString().Split('(')[1].Split(' ')[0].Replace(',', '') / 1MB), 2) } }
    }
    catch
    {
        Write-Error -Message "Error getting mailbox statistics for $($mailbox.PrimarySmtpAddress)" -ErrorAction Continue
    }

    # If an Archive exists, then get Statistics
    if ($mailbox.ArchiveStatus -ne 'None')
    {
        Write-Verbose "Getting archive mailbox statistics for $($mailbox.PrimarySmtpAddress)"
        try
        {
            $archiveStats = Get-EXOMailboxStatistics -Identity $mailbox.Guid -IncludeSoftDeletedRecipients -Properties LastLogonTime -Archive -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
            $archiveTotalItemSizeMB = $archiveStats | Select-Object @{name = 'TotalItemSizeMB'; expression = { [math]::Round(($_.TotalItemSize.ToString().Split('(')[1].Split(' ')[0].Replace(',', '') / 1MB), 2) } }
        }
        catch
        {
            Write-Error -Message "Error getting archive mailbox statistics for $($mailbox.PrimarySmtpAddress)" -ErrorAction Continue
            
        }
    }

    # Store everything in an Arraylist
    $mailboxInfo = [ordered]@{
        UserPrincipalName     = $mailbox.UserPrincipalName
        DisplayName           = $mailbox.Displayname
        PrimarySmtpAddress    = $mailbox.PrimarySmtpAddress
        Alias                 = $mailbox.Alias
        RecipientTypeDetails  = $mailbox.RecipientTypeDetails
        IsInactiveMailbox     = $mailbox.IsInactiveMailbox
        LitigationHoldEnabled = $mailbox.LitigationHoldEnabled
        RetentionHoldEnabled  = $mailbox.RetentionHoldEnabled
        InPlaceHolds          = $mailbox.InPlaceHolds -join ";"
        ArchiveStatus         = $mailbox.ArchiveStatus
    }

    if ($primaryStats)
    {
        $mailboxInfo['TotalItemSize(MB)'] = $primaryTotalItemSizeMB.TotalItemSizeMB
        $mailboxInfo['ItemCount'] = $primaryStats.ItemCount
        $mailboxInfo['DeletedItemCount'] = $primaryStats.DeletedItemCount
        $mailboxInfo['LastLogonTime'] = $primaryStats.LastLogonTime
    }
    else
    {
        $mailboxInfo['TotalItemSize(MB)'] = $null
        $mailboxInfo['ItemCount'] = $null
        $mailboxInfo['DeletedItemCount'] = $null
        $mailboxInfo['LastLogonTime'] = $null
    }

    if ($archiveStats)
    {
        $mailboxInfo['Archive_TotalItemSize(MB)'] = $archiveTotalItemSizeMB.TotalItemSizeMB
        $mailboxInfo['Archive_ItemCount'] = $archiveStats.ItemCount
        $mailboxInfo['Archive_DeletedItemCount'] = $archiveStats.DeletedItemCount
        $mailboxInfo['Archive_LastLogonTime'] = $archiveStats.LastLogonTime
    }
    else
    {
        $mailboxInfo['Archive_TotalItemSize(MB)'] = $null
        $mailboxInfo['Archive_ItemCount'] = $null
        $mailboxInfo['Archive_DeletedItemCount'] = $null
        $mailboxInfo['Archive_LastLogonTime'] = $null
    }

    Write-Verbose "Completed gathering mailbox statistics for $($mailbox.PrimarySmtpAddress)"
    return [PSCustomObject]$mailboxInfo
} #End Function Get-MailboxInformation

### Main Script
# Check if there is an active Exchange Online PowerShell session and connect if not
$PSSessions = Get-PSSession | Select-Object -Property State, Name
if ((@($PSSessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -eq 0)
{
    Write-Verbose 'Not connected to Exchange Online, prompting to connect'
    Connect-ExchangeOnline
}

# Define constants for use later
$i = 1
$timeStamp = Get-Date -Format ddMMyyyy-HHmm
Write-Verbose 'Getting Tenant Name for file name from Exchange Online'
$tenantName = (Get-OrganizationConfig).Name.Split('.')[0]
$outputFile = $OutputPath.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar + 'EXOMailboxSizes_' + $tenantName + '_' + $timeStamp + '.csv'
$output = New-Object System.Collections.Generic.List[System.Object]

Write-Verbose "Checking if $outputFile already exists"
if (Test-Path $outputFile -ErrorAction SilentlyContinue)
{
    throw "The file $outputFile already exists, please delete the file and try again."
}

# Define a hashtable for splatting into Get-EXOMailbox
$commandHashTable = @{
    Properties  = 'LitigationHoldEnabled', 'RetentionHoldEnabled', 'InPlaceHolds', 'ArchiveStatus', 'IsInactiveMailbox'
    ResultSize  = 'Unlimited'
    ErrorAction = 'Stop'
}

if ($IncludeInactiveMailboxes)
{
    $commandHashTable['IncludeInactiveMailbox'] = $true
}

if ($RecipientTypeDetails)
{
    $commandHashTable['RecipientTypeDetails'] = $RecipientTypeDetails
}

if ($MailboxFilter)
{
    $commandHashTable['Filter'] = $MailboxFilter
}

# Get mailboxes using the parameters defined from the hashtable and throw an error if encountered
try
{
    Write-Verbose 'Getting Mailboxes from Exchange Online'
    $mailboxes = @(Get-EXOMailbox @commandHashTable)
}
catch
{
    throw
}

$mailboxCount = $mailboxes.Count
Write-Verbose "There are $mailboxCount mailboxes"

if ($mailboxCount -eq 0)
{
    return 'There are no mailboxes found using the supplied filters'
}

#  Loop through the list of mailboxes and output the results to the CSV
Write-Verbose 'Beginning loop through all mailboxes'
foreach ($mailbox in $mailboxes)
{
    if (!$PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent)
    {
        Write-Progress -Id 1 -Activity 'EXO Mailbox Size Report' -Status "Processing $($i) of $($mailboxCount) Mailboxes --- $($mailbox.UserPrincipalName)" -PercentComplete (($i * 100) / $mailboxCount)
    }
    
    $mailboxInfo = Get-MailboxInformation $mailbox
    $output.Add([PSCustomObject]$mailboxInfo) | Out-Null
    $i++
}

if (!$PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent)
{
    Write-Progress -Activity 'EXO Mailbox Size Report' -Id 1 -Completed
}
$output | Export-Csv $outputFile -NoClobber -NoTypeInformation -Encoding UTF8

return "Mailbox size data has been exported to $outputfile"
