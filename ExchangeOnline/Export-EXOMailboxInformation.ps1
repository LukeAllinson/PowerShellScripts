#Requires -Version 5 -Modules @{ ModuleName = "ExchangeOnlineManagement"; ModuleVersion = "2.0.5" }, @{ ModuleName = "Join-Object"; ModuleVersion = "2.0.2" }

<#
    .SYNOPSIS
        Name: Export-EXOMailboxInformation.ps1
        This gathers extended mailbox information and exports to a csv file.

    .DESCRIPTION
        This script connects to EXO and then outputs Mailbox and CAS Mailbox information to a CSV file.

    .NOTES
        Version: 0.11
        Updated: 29-09-2025 v0.11   Updated test connection method
        Updated: 17-01-2022 v0.10   Refactored to use Join-Object (https://github.com/ili101/Join-Object)
                                    Tested on ~55k mailboxes; join took 1 minute. Previous version of script (using Where-Object) took 3 days.
        Updated: 06-01-2022 v0.9    Added CAS Mailbox information
                                    Changed output file date to match order of ISO8601 standard
        Updated: 10-11-2021 v0.8    Added parameter sets to prevent use of mutually exclusive parameters
        Updated: 10-11-2021 v0.7    Updated to include inactive mailboxes
        Updated: 08-11-2021 v0.6    Updated filename ordering
        Updated: 18-10-2021 v0.5    Refactored to simplify, improved formatting
        Updated: 15-10-2021 v0.4    Added verbose logging
        Updated: 15-10-2021 v0.3    Refactored to include error handling, filtering parameters and expanded help
        Updated: 14-10-2021 v0.2    Rewritten to improve speed and include error handling
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

    .PARAMETER Filter
        Alias of MailboxFilter parameter.

    .EXAMPLE
        .\Export-EXOMailboxInformation.ps1 C:\Scripts\ -IncludeCustomAttributes
        Exports all mailbox types including custom and extension attributes

    .EXAMPLE
        .\Export-EXOMailboxInformation.ps1 -RecipientTypeDetails RoomMailbox,EquipmentMailbox -Output C:\Scripts\
        Exports only Room and Equipment mailboxes without custom and extension attributes

    .EXAMPLE
        .\Export-EXOMailboxInformation.ps1 C:\Scripts\ -IncludeCustomAttributes -MailboxFilter 'Department -eq "R&D"'
        Exports all mailboxes from the R&D department with custom and extension attributes
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
        ParameterSetName = 'DefaultParameters'
    )]
    [Parameter(
        ParameterSetName = 'InactiveOnly'
    )]
    [Parameter(
        ParameterSetName = 'IncludeInactive'
    )]
    [switch]
    $IncludeCustomAttributes,
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

### Main Script
# Check if there is an active Exchange Online PowerShell session and connect if not
$exoSessions = Get-ConnectionInformation | Select-Object -Property Name, State
if (($exoSessions | Where-Object { ($_.State -eq 'Connected') -and ($_.Name -like 'ExchangeOnline_*') }).Count -eq 0)
{
    Write-Verbose 'Not connected to Exchange Online, prompting to connect'
    Connect-ExchangeOnline
}

# Define constants for use later
$timeStamp = Get-Date -Format yyyyMMdd-HHmm
Write-Verbose 'Getting Tenant Name for file name from Exchange Online'
$tenantName = (Get-OrganizationConfig).Name.Split('.')[0]
$outputFile = $OutputPath.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar + 'EXOMailboxInformation_' + $tenantName + '_' + $timeStamp + '.csv'

Write-Verbose "Checking if $outputFile already exists"
if (Test-Path $outputFile -ErrorAction SilentlyContinue)
{
    throw "The file $outputFile already exists, please delete the file and try again."
}

# Define a hashtable for splatting into Get-EXOMailbox
$commandHashTable = @{
    ResultSize  = 'Unlimited'
    ErrorAction = 'Stop'
}

if ($InactiveMailboxOnly)
{
    $commandHashTable['InactiveMailboxOnly'] = $true
}

if ($IncludeInactiveMailboxes)
{
    $commandHashTable['IncludeInactiveMailbox'] = $true
}

if (!$IncludeCustomAttributes)
{
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
    $commandHashTable['Properties'] = $mailboxProperties
}
else
{
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
    'UsageLocation',
    'CustomAttribute1',
    'CustomAttribute2',
    'CustomAttribute3',
    'CustomAttribute4',
    'CustomAttribute5',
    'CustomAttribute6',
    'CustomAttribute7',
    'CustomAttribute8',
    'CustomAttribute9',
    'CustomAttribute10',
    'CustomAttribute11',
    'CustomAttribute12',
    'CustomAttribute13',
    'CustomAttribute14',
    'CustomAttribute15',
    'ExtensionCustomAttribute1',
    'ExtensionCustomAttribute2',
    'ExtensionCustomAttribute3',
    'ExtensionCustomAttribute4',
    'ExtensionCustomAttribute5'
    $commandHashTable['Properties'] = $mailboxProperties
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
    $mailboxes = @(Get-EXOMailbox @commandHashTable | Select-Object $mailboxProperties)
}
catch
{
    throw
}

Write-Verbose "There are $mailboxCount mailboxes"
$mailboxCount = $mailboxes.Count

if ($mailboxCount -eq 0)
{
    return 'There are no mailboxes found using the supplied filters'
}

# Define a hashtable for splatting into Get-EXOMailbox
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

$casCommandHashTable = @{
    ResultSize  = 'Unlimited'
    ErrorAction = 'Stop'
    Properties  = $casMailboxProperties
}

if ($MailboxFilter)
{
    $casCommandHashTable['Filter'] = $MailboxFilter
}

# Get casMailboxes using the parameters defined from the hashtable and throw an error if encountered
try
{
    Write-Verbose 'Getting CAS Mailboxes from Exchange Online'
    $casMailboxes = @(Get-EXOCasMailbox @casCommandHashTable | Select-Object $casMailboxProperties)
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
