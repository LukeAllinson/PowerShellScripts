#Requires -Version 5 -Modules ExchangeOnlineManagement

<#
    .SYNOPSIS
        Name: Export-EXOMailboxInformation.ps1
        This gathers extended mailbox information and exports to a csv file.

    .DESCRIPTION
        This script connects to EXO and then outputs Mailbox information to a CSV file.

    .NOTES
        Version: 0.9
        TODO: Add location information --- Added usage location but nothing else really available
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

function Get-ExoCasMailboxInfo
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [System.Guid]
        $Guid,
        [Parameter(Mandatory)]
        [array]
        $CasMailboxes
    )

    return $CasMailboxes.where({ $_.Guid -eq $Guid })
}

### Main Script
# Check if there is an active Exchange Online PowerShell session and connect if not
$PSSessions = Get-PSSession | Select-Object -Property State, Name
if ((@($PSSessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -eq 0)
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
    $commandHashTable['Properties'] = 'UserPrincipalName',
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
}
else
{
    $commandHashTable['Properties'] = 'UserPrincipalName',
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

Write-Verbose "There are $mailboxCount mailboxes"
$mailboxCount = $mailboxes.Count

if ($mailboxCount -eq 0)
{
    return 'There are no mailboxes found using the supplied filters'
}

$casCommandHashTable = @{
    ResultSize  = 'Unlimited'
    ErrorAction = 'Stop'
    Properties  = 'Guid',
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
}

if ($MailboxFilter)
{
    $casCommandHashTable['Filter'] = $MailboxFilter
}

try
{
    Write-Verbose 'Getting CAS Mailboxes from Exchange Online'
    $casMailboxes = @(Get-EXOCasMailbox @casCommandHashTable)
}
catch
{
    throw
}

# Select the required properties and export to csv
Write-Verbose 'Processing data'
if (!$IncludeCustomAttributes)
{
    $mailboxes | Select-Object -Property 'UserPrincipalName',
    'Name',
    'DisplayName',
    'SimpleDisplayName',
    'PrimarySmtpAddress',
    'Alias',
    'SamAccountName',
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
    'UsageLocation',
    'Guid',
    'ExchangeGuid',
    'ArchiveStatus',
    'ArchiveName',
    'ArchiveGuid',
    @{ Name = 'EmailAddresses'; Expression = { $($_.EmailAddresses -join ';') } },
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
    'MaxSendSize',
    'MaxReceiveSize',
    @{ Name = 'UniversalOutlookEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).UniversalOutlookEnabled) } },
    @{ Name = 'OutlookMobileEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).OutlookMobileEnabled) } },
    @{ Name = 'MacOutlookEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).MacOutlookEnabled) } },
    @{ Name = 'ECPEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ECPEnabled) } },
    @{ Name = 'OWAforDevicesEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).OWAforDevicesEnabled) } },
    @{ Name = 'ShowGalAsDefaultView'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ShowGalAsDefaultView) } },
    @{ Name = 'SmtpClientAuthenticationDisabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).SmtpClientAuthenticationDisabled) } },
    @{ Name = 'OWAEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).OWAEnabled) } },
    @{ Name = 'PublicFolderClientAccess'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).PublicFolderClientAccess) } },
    @{ Name = 'OwaMailboxPolicy'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).OwaMailboxPolicy) } },
    @{ Name = 'IsOptimizedForAccessibility'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).IsOptimizedForAccessibility) } },
    @{ Name = 'ImapEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ImapEnabled) } },
    @{ Name = 'ImapSuppressReadReceipt'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ImapSuppressReadReceipt) } },
    @{ Name = 'ImapEnableExactRFC822Size'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ImapEnableExactRFC822Size) } },
    @{ Name = 'ImapMessagesRetrievalMimeFormat'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ImapMessagesRetrievalMimeFormat) } },
    @{ Name = 'ImapUseProtocolDefaults'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ImapUseProtocolDefaults) } },
    @{ Name = 'ImapForceICalForCalendarRetrievalOption'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ImapForceICalForCalendarRetrievalOption) } },
    @{ Name = 'PopEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).PopEnabled) } },
    @{ Name = 'PopSuppressReadReceipt'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).PopSuppressReadReceipt) } },
    @{ Name = 'PopEnableExactRFC822Size'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).PopEnableExactRFC822Size) } },
    @{ Name = 'PopMessagesRetrievalMimeFormat'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).PopMessagesRetrievalMimeFormat) } },
    @{ Name = 'PopUseProtocolDefaults'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).PopUseProtocolDefaults) } },
    @{ Name = 'PopMessageDeleteEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).PopMessageDeleteEnabled) } },
    @{ Name = 'PopForceICalForCalendarRetrievalOption'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).PopForceICalForCalendarRetrievalOption) } },
    @{ Name = 'MAPIEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).MAPIEnabled) } },
    @{ Name = 'MAPIBlockOutlookVersions'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).MAPIBlockOutlookVersions) } },
    @{ Name = 'MAPIBlockOutlookRpcHttp'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).MAPIBlockOutlookRpcHttp) } },
    @{ Name = 'MapiHttpEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).MapiHttpEnabled) } },
    @{ Name = 'MAPIBlockOutlookNonCachedMode'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).MAPIBlockOutlookNonCachedMode) } },
    @{ Name = 'MAPIBlockOutlookExternalConnectivity'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).MAPIBlockOutlookExternalConnectivity) } },
    @{ Name = 'EwsEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).EwsEnabled) } },
    @{ Name = 'EwsAllowOutlook'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).EwsAllowOutlook) } },
    @{ Name = 'EwsAllowMacOutlook'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).EwsAllowMacOutlook) } },
    @{ Name = 'EwsAllowEntourage'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).EwsAllowEntourage) } },
    @{ Name = 'EwsApplicationAccessPolicy'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).EwsApplicationAccessPolicy) } },
    @{ Name = 'EwsAllowList'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).EwsAllowList) } },
    @{ Name = 'EwsBlockList'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).EwsBlockList) } },
    @{ Name = 'ActiveSyncAllowedDeviceIDs'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ActiveSyncAllowedDeviceIDs) } },
    @{ Name = 'ActiveSyncBlockedDeviceIDs'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ActiveSyncBlockedDeviceIDs) } },
    @{ Name = 'ActiveSyncEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ActiveSyncEnabled) } },
    @{ Name = 'ActiveSyncSuppressReadReceipt'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ActiveSyncSuppressReadReceipt) } },
    @{ Name = 'ActiveSyncMailboxPolicyIsDefaulted'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ActiveSyncMailboxPolicyIsDefaulted) } },
    @{ Name = 'ActiveSyncMailboxPolicy'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ActiveSyncMailboxPolicy) } },
    @{ Name = 'HasActiveSyncDevicePartnership'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).HasActiveSyncDevicePartnership) } },
    @{ Name = 'ActiveSyncDebugLogging'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ActiveSyncDebugLogging) } } | Export-Csv $outputFile -NoClobber -NoTypeInformation -Encoding UTF8
}
else
{
    $mailboxes | Select-Object -Property 'UserPrincipalName',
    'Name',
    'DisplayName',
    'SimpleDisplayName',
    'PrimarySmtpAddress',
    'Alias',
    'SamAccountName',
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
    'UsageLocation',
    'Guid',
    'ExchangeGuid',
    'ArchiveStatus',
    'ArchiveName',
    'ArchiveGuid',
    @{ Name = 'EmailAddresses'; Expression = { $($_.EmailAddresses -join ';') } },
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
    'MaxSendSize',
    'MaxReceiveSize',
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
    'ExtensionCustomAttribute5',
    @{ Name = 'UniversalOutlookEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).UniversalOutlookEnabled) } },
    @{ Name = 'OutlookMobileEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).OutlookMobileEnabled) } },
    @{ Name = 'MacOutlookEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).MacOutlookEnabled) } },
    @{ Name = 'ECPEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ECPEnabled) } },
    @{ Name = 'OWAforDevicesEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).OWAforDevicesEnabled) } },
    @{ Name = 'ShowGalAsDefaultView'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ShowGalAsDefaultView) } },
    @{ Name = 'SmtpClientAuthenticationDisabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).SmtpClientAuthenticationDisabled) } },
    @{ Name = 'OWAEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).OWAEnabled) } },
    @{ Name = 'PublicFolderClientAccess'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).PublicFolderClientAccess) } },
    @{ Name = 'OwaMailboxPolicy'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).OwaMailboxPolicy) } },
    @{ Name = 'IsOptimizedForAccessibility'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).IsOptimizedForAccessibility) } },
    @{ Name = 'ImapEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ImapEnabled) } },
    @{ Name = 'ImapSuppressReadReceipt'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ImapSuppressReadReceipt) } },
    @{ Name = 'ImapEnableExactRFC822Size'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ImapEnableExactRFC822Size) } },
    @{ Name = 'ImapMessagesRetrievalMimeFormat'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ImapMessagesRetrievalMimeFormat) } },
    @{ Name = 'ImapUseProtocolDefaults'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ImapUseProtocolDefaults) } },
    @{ Name = 'ImapForceICalForCalendarRetrievalOption'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ImapForceICalForCalendarRetrievalOption) } },
    @{ Name = 'PopEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).PopEnabled) } },
    @{ Name = 'PopSuppressReadReceipt'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).PopSuppressReadReceipt) } },
    @{ Name = 'PopEnableExactRFC822Size'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).PopEnableExactRFC822Size) } },
    @{ Name = 'PopMessagesRetrievalMimeFormat'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).PopMessagesRetrievalMimeFormat) } },
    @{ Name = 'PopUseProtocolDefaults'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).PopUseProtocolDefaults) } },
    @{ Name = 'PopMessageDeleteEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).PopMessageDeleteEnabled) } },
    @{ Name = 'PopForceICalForCalendarRetrievalOption'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).PopForceICalForCalendarRetrievalOption) } },
    @{ Name = 'MAPIEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).MAPIEnabled) } },
    @{ Name = 'MAPIBlockOutlookVersions'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).MAPIBlockOutlookVersions) } },
    @{ Name = 'MAPIBlockOutlookRpcHttp'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).MAPIBlockOutlookRpcHttp) } },
    @{ Name = 'MapiHttpEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).MapiHttpEnabled) } },
    @{ Name = 'MAPIBlockOutlookNonCachedMode'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).MAPIBlockOutlookNonCachedMode) } },
    @{ Name = 'MAPIBlockOutlookExternalConnectivity'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).MAPIBlockOutlookExternalConnectivity) } },
    @{ Name = 'EwsEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).EwsEnabled) } },
    @{ Name = 'EwsAllowOutlook'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).EwsAllowOutlook) } },
    @{ Name = 'EwsAllowMacOutlook'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).EwsAllowMacOutlook) } },
    @{ Name = 'EwsAllowEntourage'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).EwsAllowEntourage) } },
    @{ Name = 'EwsApplicationAccessPolicy'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).EwsApplicationAccessPolicy) } },
    @{ Name = 'EwsAllowList'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).EwsAllowList) } },
    @{ Name = 'EwsBlockList'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).EwsBlockList) } },
    @{ Name = 'ActiveSyncAllowedDeviceIDs'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ActiveSyncAllowedDeviceIDs) } },
    @{ Name = 'ActiveSyncBlockedDeviceIDs'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ActiveSyncBlockedDeviceIDs) } },
    @{ Name = 'ActiveSyncEnabled'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ActiveSyncEnabled) } },
    @{ Name = 'ActiveSyncSuppressReadReceipt'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ActiveSyncSuppressReadReceipt) } },
    @{ Name = 'ActiveSyncMailboxPolicyIsDefaulted'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ActiveSyncMailboxPolicyIsDefaulted) } },
    @{ Name = 'ActiveSyncMailboxPolicy'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ActiveSyncMailboxPolicy) } },
    @{ Name = 'HasActiveSyncDevicePartnership'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).HasActiveSyncDevicePartnership) } },
    @{ Name = 'ActiveSyncDebugLogging'; Expression = { $((Get-ExoCasMailboxInfo -Guid $_.Guid -CasMailboxes $casMailboxes).ActiveSyncDebugLogging) } } | Export-Csv $outputFile -NoClobber -NoTypeInformation -Encoding UTF8
}

return "Mailbox information has been exported to $outputfile"
