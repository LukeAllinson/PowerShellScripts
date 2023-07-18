#Requires -Version 5 -Modules @{ ModuleName = "Join-Object"; ModuleVersion = "2.0.2" }

<#
    .SYNOPSIS
        Name: Export-MailboxInformation.ps1
        This gathers extended mailbox information and exports to a csv file.

    .DESCRIPTION
        This script outputs Mailbox and CAS Mailbox information to a CSV file.

    .NOTES
        Authors: Luke Allinson (github:LukeAllinson)
                 Robin Dadswell (github:RobinDadswell)

    .PARAMETER OutputPath
        Full path to the folder where the output will be saved.
        Can be used without the parameter name in the first position only.

    .PARAMETER InactiveMailboxOnly
        Only gathers information about inactive mailboxes (active mailboxes are not included in results).

    .PARAMETER IncludeInactiveMailboxes
        Include inactive mailboxes in results; these are not included by default.

    .PARAMETER IncludeCustomAttributes
        Include custom and extension attributes; these are not included by default.

    .PARAMETER RecipientTypeDetails
        Provide one or more RecipientTypeDetails values to return only mailboxes of those types in the results. Seperate multiple values by commas.
        Valid values are: DiscoveryMailbox, EquipmentMailbox, GroupMailbox, RoomMailbox, SchedulingMailbox, SharedMailbox, TeamMailbox, UserMailbox.

    .PARAMETER MailboxFilter
        Provide a filter to reduce the size of the Get-Mailbox query; this must follow oPath syntax standards.
        For example:
        'EmailAddresses -like "*bruce*"'
        'DisplayName -like "*wayne*"'
        'CustomAttribute1 -eq "InScope"'

    .PARAMETER Filter
        Alias of MailboxFilter parameter.

    .EXAMPLE
        .\Export-MailboxInformation.ps1 C:\Scripts\ -IncludeCustomAttributes
        Exports all mailbox types including custom and extension attributes

    .EXAMPLE
        .\Export-MailboxInformation.ps1 -RecipientTypeDetails RoomMailbox,EquipmentMailbox -Output C:\Scripts\
        Exports only Room and Equipment mailboxes without custom and extension attributes

    .EXAMPLE
        .\Export-MailboxInformation.ps1 C:\Scripts\ -IncludeCustomAttributes -MailboxFilter 'Department -eq "R&D"'
        Exports all mailboxes from the R&D department with custom and extension attributes
#>

[CmdletBinding(DefaultParameterSetName = 'DefaultParameters')]
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
    [Parameter()]
    [Alias('Filter')]
    [string]
    $MailboxFilter
)

### Main Script
# Define constants for use later
$timeStamp = Get-Date -Format yyyyMMdd-HHmm
Write-Verbose 'Getting Tenant Name for file name from Exchange'
$outputFile = $OutputPath.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar + 'MailboxInformation_' + $timeStamp + '.csv'

Write-Verbose "Checking if $outputFile already exists"
if (Test-Path $outputFile -ErrorAction SilentlyContinue)
{
    throw "The file $outputFile already exists, please delete the file and try again."
}

# Define a hashtable for splatting into Get-Mailbox
$commandHashTable = @{
    ResultSize  = 'Unlimited'
    ErrorAction = 'Stop'
}

if ($RecipientTypeDetails)
{
    $commandHashTable['RecipientTypeDetails'] = $RecipientTypeDetails
}

if ($MailboxFilter)
{
    $commandHashTable['Filter'] = $MailboxFilter
}

[System.Collections.ArrayList]$mailboxProperties = 'UserPrincipalName',
'Name',
'DisplayName',
'SimpleDisplayName',
'PrimarySmtpAddress',
'Alias',
'SamAccountName',
'ExchangeGuid',
'Guid',
'RecipientTypeDetails',
'Database',
'ForwardingAddress',
'ForwardingSmtpAddress',
'DeliverToMailboxAndForward',
'LitigationHoldEnabled',
'RetentionHoldEnabled',
'InPlaceHolds',
'RetentionPolicy',
'IsInactiveMailbox',
'InactiveMailboxRetireTime',
'HiddenFromAddressListsEnabled',
'ArchiveStatus',
'ArchiveName',
'ArchiveGuid',
'EmailAddresses',
'WhenChanged',
'WhenChangedUTC',
'WhenMailboxCreated',
'WhenCreated',
'WhenCreatedUTC',
'UMEnabled',
'ExternalOofOptions',
'IssueWarningQuota',
'ProhibitSendQuota',
'ProhibitSendReceiveQuota',
'UseDatabaseQuotaDefaults',
'MaxSendSize',
'MaxReceiveSize',
'UsageLocation'

# Get mailboxes using the parameters defined from the hashtable and throw an error if encountered
try
{
    Write-Verbose 'Getting Mailboxes from Exchange'
    $mailboxes = @(Get-Mailbox @commandHashTable | Select-Object $mailboxProperties)
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

$casCommandHashTable = @{
    ResultSize  = 'Unlimited'
    ErrorAction = 'Stop'
}


if ($MailboxFilter)
{
    $casCommandHashTable['Filter'] = $MailboxFilter
}

[System.Collections.ArrayList]$casMailboxProperties = 'Guid',
'UniversalOutlookEnabled',
'OutlookMobileEnabled',
'MacOutlookEnabled',
'ECPEnabled',
'OWAforDevicesEnabled',
'ShowGalAsDefaultView',
'SmtpClientAuthenticationDisabled',
'OWAEnabled',
'PublicFolderClientAccess',
'OwaMailboxPolicy',
'ImapEnabled',
'ImapSuppressReadReceipt',
'ImapEnableExactRFC822Size',
'ImapMessagesRetrievalMimeFormat',
'ImapUseProtocolDefaults',
'ImapForceICalForCalendarRetrievalOption',
'PopEnabled',
'PopSuppressReadReceipt',
'PopEnableExactRFC822Size',
'PopMessagesRetrievalMimeFormat',
'PopUseProtocolDefaults',
'PopMessageDeleteEnabled',
'PopForceICalForCalendarRetrievalOption',
'MAPIEnabled',
'MAPIBlockOutlookVersions',
'MAPIBlockOutlookRpcHttp',
'MapiHttpEnabled',
'MAPIBlockOutlookNonCachedMode',
'MAPIBlockOutlookExternalConnectivity',
'EwsEnabled',
'EwsAllowOutlook',
'EwsAllowMacOutlook',
'EwsAllowEntourage',
'EwsApplicationAccessPolicy',
'EwsAllowList',
'EwsBlockList',
'ActiveSyncAllowedDeviceIDs',
'ActiveSyncBlockedDeviceIDs',
'ActiveSyncEnabled',
'ActiveSyncSuppressReadReceipt',
'ActiveSyncMailboxPolicyIsDefaulted',
'ActiveSyncMailboxPolicy',
'HasActiveSyncDevicePartnership'

# Get casMailboxes using the parameters defined from the hashtable and throw an error if encountered
try
{
    Write-Verbose 'Getting CAS Mailboxes from Exchange'
    $casMailboxes = @(Get-CasMailbox @casCommandHashTable | Select-Object $casMailboxProperties)
}
catch
{
    throw
}

Write-Verbose 'Joining mailbox and casMailbox data'
# start a stopwatch to time the join process
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# join the $mailboxes and $casMailboxes arrayLists
$combinedMailboxes = Join-Object -Left $mailboxes -Right $casMailboxes -LeftJoinProperty Guid -RightJoinProperty Guid -Type OnlyIfInBoth

$stopwatch.Stop()

# display time taken to join
Write-Verbose "Time taken to join: $($stopwatch.Elapsed)"

# set up combined properties array for Select-Object
$combinedMailboxProperties = [System.Collections.ArrayList]@()
$combinedMailboxProperties.AddRange($mailboxProperties)
$combinedMailboxProperties.AddRange($casMailboxProperties)
$combinedMailboxProperties = $combinedMailboxProperties | Select-Object -Unique

# export combined results to CSV
Write-Verbose 'Writing CSV file'
$combinedMailboxes | Select-Object $combinedMailboxProperties | Export-Csv $outputFile -NoClobber -NoTypeInformation -Encoding UTF8

return "Mailbox information has been exported to $outputfile"
