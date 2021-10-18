#Requires -Version 5 -Modules ExchangeOnlineManagement

<#
	.SYNOPSIS
		Name: Get-EXOMailboxInformation.ps1
		This gathers extended mailbox information.

	.DESCRIPTION
		This script connects to EXO and then outputs Mailbox information to a CSV file.

	.NOTES
		Version: 0.5
        Updated: 18-10-2021 v0.5    Refactored to simplify
        Updated: 15-10-2021 v0.4    Added verbose logging
        Updated: 15-10-2021 v0.3    Refactored to include error handling, filtering parameters and expanded help
        Updated: 14-10-2021 v0.2    Rewritten to improve speed and include error handling
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
    [switch]
    $IncludeCustomAttributes,
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
    $mailboxInfo = [ordered]@{
        UserPrincipalName = $mailbox.UserPrincipalName
        Name = $mailbox.Name
        DisplayName = $mailbox.Displayname
        SimpleDisplayName = $mailbox.SimpleDisplayName
        PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
        Alias = $mailbox.Alias
        SamAccountName = $mailbox.SamAccountName
        RecipientTypeDetails = $mailbox.RecipientTypeDetails
        ForwardingAddress = $mailbox.ForwardingAddress
        ForwardingSmtpAddress = $mailbox.ForwardingSmtpAddress
        DeliverToMailboxAndForward = $mailbox.DeliverToMailboxAndForward
        LitigationHoldEnabled = $mailbox.LitigationHoldEnabled
        RetentionHoldEnabled = $mailbox.RetentionHoldEnabled
        InPlaceHolds = $mailbox.InPlaceHolds
        GrantSendOnBehalfTo = $mailbox.GrantSendOnBehalfTo
        HiddenFromAddressListsEnabled = $mailbox.HiddenFromAddressListsEnabled
        ExchangeGuid = $mailbox.ExchangeGuid
        ArchiveStatus = $mailbox.ArchiveStatus
        ArchiveName = $mailbox.ArchiveName
        ArchiveGuid = $mailbox.ArchiveGuid
        EmailAddresses = ($mailbox.EmailAddresses -join ";")
        WhenChanged = $mailbox.WhenChanged
        WhenChangedUTC = $mailbox.WhenChangedUTC
        WhenMailboxCreated = $mailbox.WhenMailboxCreated
        WhenCreated = $mailbox.WhenCreated
        WhenCreatedUTC = $mailbox.WhenCreatedUTC
        UMEnabled = $mailbox.UMEnabled
        ExternalOofOptions = $mailbox.ExternalOofOptions
        IssueWarningQuota = $mailbox.IssueWarningQuota
        ProhibitSendQuota = $mailbox.ProhibitSendQuota
        ProhibitSendReceiveQuota = $mailbox.ProhibitSendReceiveQuota
        UseDatabaseQuotaDefaults = $mailbox.UseDatabaseQuotaDefaults
        MaxSendSize = $mailbox.MaxSendSize
        MaxReceiveSize = $mailbox.MaxReceiveSize
    }

    if ($IncludeCustomAttributes) {
        $mailboxInfo["CustomAttribute1"] = $mailbox.CustomAttribute1
        $mailboxInfo["CustomAttribute2"] = $mailbox.CustomAttribute2
        $mailboxInfo["CustomAttribute3"] = $mailbox.CustomAttribute3
        $mailboxInfo["CustomAttribute4"] = $mailbox.CustomAttribute4
        $mailboxInfo["CustomAttribute5"] = $mailbox.CustomAttribute5
        $mailboxInfo["CustomAttribute6"] = $mailbox.CustomAttribute6
        $mailboxInfo["CustomAttribute7"] = $mailbox.CustomAttribute7
        $mailboxInfo["CustomAttribute8"] = $mailbox.CustomAttribute8
        $mailboxInfo["CustomAttribute9"] = $mailbox.CustomAttribute9
        $mailboxInfo["CustomAttribute10"] = $mailbox.CustomAttribute10
        $mailboxInfo["CustomAttribute11"] = $mailbox.CustomAttribute11
        $mailboxInfo["CustomAttribute12"] = $mailbox.CustomAttribute12
        $mailboxInfo["CustomAttribute13"] = $mailbox.CustomAttribute13
        $mailboxInfo["CustomAttribute14"] = $mailbox.CustomAttribute14
        $mailboxInfo["CustomAttribute15"] = $mailbox.CustomAttribute15
        $mailboxInfo["ExtensionCustomAttribute1"] = $mailbox.ExtensionCustomAttribute1
        $mailboxInfo["ExtensionCustomAttribute2"] = $mailbox.ExtensionCustomAttribute2
        $mailboxInfo["ExtensionCustomAttribute3"] = $mailbox.ExtensionCustomAttribute3
        $mailboxInfo["ExtensionCustomAttribute4"] = $mailbox.ExtensionCustomAttribute4
        $mailboxInfo["ExtensionCustomAttribute5"] = $mailbox.ExtensionCustomAttribute5
    }
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
$timeStamp = Get-Date -Format ddMMyyyy-HHmm
Write-Verbose "Getting Tenant Name for file name from Exchange Online"
$tenantName = (Get-OrganizationConfig).Name.Split(".")[0]
$outputFile = $OutputPath.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar + $timeStamp + '-' + $tenantName + '-' + 'EXOMailboxInformation.csv'

Write-Verbose "Checking if $outputFile already exists"
if (Test-Path $outputFile -ErrorAction SilentlyContinue) {
    throw "The file $outputFile already exists, please delete the file and try again."
}

# Define a hashtable for splatting into Get-EXOMailbox
$commandHashTable = @{
    ResultSize = "Unlimited"
    ErrorAction = "Stop"
}

if (!$IncludeCustomAttributes) {
    $commandHashTable["Properties"] = "UserPrincipalName","Name","DisplayName","SimpleDisplayName","PrimarySmtpAddress","Alias","SamAccountName","ExchangeGuid","RecipientTypeDetails","ForwardingAddress","ForwardingSmtpAddress","DeliverToMailboxAndForward","LitigationHoldEnabled","RetentionHoldEnabled","InPlaceHolds","GrantSendOnBehalfTo","HiddenFromAddressListsEnabled","ArchiveStatus","ArchiveName","ArchiveGuid","EmailAddresses","WhenChanged","WhenChangedUTC","WhenMailboxCreated","WhenCreated","WhenCreatedUTC","UMEnabled","ExternalOofOptions","IssueWarningQuota","ProhibitSendQuota","ProhibitSendReceiveQuota","UseDatabaseQuotaDefaults","MaxSendSize","MaxReceiveSize"
} else {
    $commandHashTable["Properties"] = "UserPrincipalName","Name","DisplayName","SimpleDisplayName","PrimarySmtpAddress","Alias","SamAccountName","ExchangeGuid","RecipientTypeDetails","ForwardingAddress","ForwardingSmtpAddress","DeliverToMailboxAndForward","LitigationHoldEnabled","RetentionHoldEnabled","InPlaceHolds","GrantSendOnBehalfTo","HiddenFromAddressListsEnabled","ArchiveStatus","ArchiveName","ArchiveGuid","EmailAddresses","WhenChanged","WhenChangedUTC","WhenMailboxCreated","WhenCreated","WhenCreatedUTC","UMEnabled","ExternalOofOptions","IssueWarningQuota","ProhibitSendQuota","ProhibitSendReceiveQuota","UseDatabaseQuotaDefaults","MaxSendSize","MaxReceiveSize","CustomAttribute1","CustomAttribute2","CustomAttribute3","CustomAttribute4","CustomAttribute5","CustomAttribute6","CustomAttribute7","CustomAttribute8","CustomAttribute9","CustomAttribute10","CustomAttribute11","CustomAttribute12","CustomAttribute13","CustomAttribute14","CustomAttribute15","ExtensionCustomAttribute1","ExtensionCustomAttribute2","ExtensionCustomAttribute3","ExtensionCustomAttribute4","ExtensionCustomAttribute5"
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

if ($mailboxCount -eq 0) {
    throw "There are no mailboxes found using the filters requested."
}

if (!$IncludeCustomAttributes) {
    $mailboxes | Select-Object -Property "UserPrincipalName",
        "Name",
        "DisplayName",
        "SimpleDisplayName",
        "PrimarySmtpAddress",
        "Alias",
        "SamAccountName",
        "RecipientTypeDetails",
        "ForwardingAddress",
        "ForwardingSmtpAddress",
        "DeliverToMailboxAndForward",
        "LitigationHoldEnabled",
        "RetentionHoldEnabled",
        "InPlaceHolds",
        "GrantSendOnBehalfTo",
        "HiddenFromAddressListsEnabled",
        "ExchangeGuid",
        "ArchiveStatus",
        "ArchiveName",
        "ArchiveGuid",
        @{ Name = 'EmailAddresses'; Expression = { $($_.EmailAddresses -join ";") } },
        "WhenChanged",
        "WhenChangedUTC",
        "WhenMailboxCreated",
        "WhenCreated",
        "WhenCreatedUTC",
        "UMEnabled",
        "ExternalOofOptions",
        "IssueWarningQuota",
        "ProhibitSendQuota",
        "ProhibitSendReceiveQuota",
        "MaxSendSize",
        "MaxReceiveSize" | Export-Csv $outputFile -NoClobber -NoTypeInformation -Encoding UTF8
} else {
    $mailboxes | Select-Object -Property "UserPrincipalName",
        "Name",
        "DisplayName",
        "SimpleDisplayName",
        "PrimarySmtpAddress",
        "Alias",
        "SamAccountName",
        "RecipientTypeDetails",
        "ForwardingAddress",
        "ForwardingSmtpAddress",
        "DeliverToMailboxAndForward",
        "LitigationHoldEnabled",
        "RetentionHoldEnabled",
        "InPlaceHolds",
        "GrantSendOnBehalfTo",
        "HiddenFromAddressListsEnabled",
        "ExchangeGuid",
        "ArchiveStatus",
        "ArchiveName",
        "ArchiveGuid",
        @{ Name = 'EmailAddresses'; Expression = { $($_.EmailAddresses -join ";") } },
        "WhenChanged",
        "WhenChangedUTC",
        "WhenMailboxCreated",
        "WhenCreated",
        "WhenCreatedUTC",
        "UMEnabled",
        "ExternalOofOptions",
        "IssueWarningQuota",
        "ProhibitSendQuota",
        "ProhibitSendReceiveQuota",
        "MaxSendSize",
        "MaxReceiveSize",
        "CustomAttribute1",
        "CustomAttribute2",
        "CustomAttribute3",
        "CustomAttribute4",
        "CustomAttribute5",
        "CustomAttribute6",
        "CustomAttribute7",
        "CustomAttribute8",
        "CustomAttribute9",
        "CustomAttribute10",
        "CustomAttribute11",
        "CustomAttribute12",
        "CustomAttribute13",
        "CustomAttribute14",
        "CustomAttribute15",
        "ExtensionCustomAttribute1",
        "ExtensionCustomAttribute2",
        "ExtensionCustomAttribute3",
        "ExtensionCustomAttribute4",
        "ExtensionCustomAttribute5" | Export-Csv $outputFile -NoClobber -NoTypeInformation -Encoding UTF8
}

return "Mailbox information has been exported to $outputfile"
