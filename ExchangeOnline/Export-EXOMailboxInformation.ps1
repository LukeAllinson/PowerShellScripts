#Requires -Version 5 -Modules ExchangeOnlineManagement

<#
	.SYNOPSIS
		Name: Get-EXOMailboxInformation.ps1
		This gathers extended mailbox information.

	.DESCRIPTION
		This script connects to EXO and then outputs Mailbox information to a CSV file. 

	.NOTES
		Version: 0.2
        Updated: 14-10-2021 v0.2    Rewritten to improve speed and include error handling
		Updated: <unknown>	v0.1	Initial draft

		Authors: Luke Allinson, Robin Dadswell
#>

param
(
    [Parameter(
        Mandatory,
        HelpMessage = "Full path to the folder where the output will be saved."
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateScript(
        {
            if (!(Test-Path -Path $_)) {
                throw "The folder $_ does not exist"
            } else {
                return $true
            }
        })]
    [IO.DirectoryInfo]
    $OutputPath,
    [Parameter(
        HelpMessage = "Do not include custom and extension attributes."
    )]
    [switch]
    $IncludeCustomAttributes,
    [Parameter(
        HelpMessage = "Use custom mailbox filter."
    )]
    [string]
    $MailboxFilter,
    [Parameter(
        HelpMessage = "Specify RecipientTypeDetails to filter results."
    )]
    [string]
    $RecipientTypeDetails
)

function Get-MailboxInformation ($mailbox) {
    # Store everything in an Arraylist
    $mailboxInfo = [System.Collections.ArrayList]@()
    $mailboxInfo = @{
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
    Return [PSCustomObject]$mailboxInfo
} #End Function Get-MailboxInformation

# Main Script
$i = 1
$timeStamp = Get-Date -Format ddMMyyyy-HHmm
$tenantName = (Get-OrganizationConfig).Name.Split(".")[0]
$outputFile = $OutputPath.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar + $timeStamp + '-' + $tenantName + '-' + 'EXOMailboxPermissions.csv'
$output = [System.Collections.ArrayList]@()

# Build Filters if required
$filters = $null
if ($MailboxFilter) {
    $filters = "-Filter $MailboxFilter"
}
if ($RecipientTypeDetails) {
    if ($filters) {
        $filters += " -RecipientTypeDetails $RecipientTypeDetails"
    } else {
        $filters += "-RecipientTypeDetails $RecipientTypeDetails"
    }
}
if ($filters) {
    $mailboxes = @(Get-EXOMailbox $filters -Resultsize Unlimited -Properties UserPrincipalName,Name,DisplayName,SimpleDisplayName,PrimarySmtpAddress,Alias,SamAccountName,ExchangeGuid,RecipientTypeDetails,ForwardingAddress,ForwardingSmtpAddress,DeliverToMilboxAndForward,LitigationHoldEnabled,RetentionHoldEnabled,InPlaceHolds,GrantSendOnBehalfTo,HiddenFromAddressListsEnabled,ArchiveStatus,ArchiveName,ArchiveGuid,EmailAddresses,WhenChanged,WhenChangedUTC,WhenMailboxCreated,WhenCreated,WhenCreatedUTC,UMEnabled,ExternalOofOptions,IssueWarningQuota,ProhiitSendQuota,ProhibitSendReceiveQuota,UseDatabaseQuotaDefaults,MaxSendSize,MaxReceiveSize,CustomAttribute1,CustomAttribute2,CustomAttribute3,CustomAttribue4,CustomAttribute5,CustomAttribute6,CustomAttribute7,CustomAttribute8,CustomAttribute9,CustomAttribute10,CustomAttribute11,CustomAttribute12,CustomAttriute13,CustomAttribute14,CustomAttribute15,ExtensionCustomAttribute1,ExtensionCustomAttribute2,ExtensionCustomAttribute3,ExtensionCustomAttribute4,ExtensionCustomAttribute5)
} else {
    $mailboxes = @(Get-EXOMailbox -Resultsize Unlimited -Properties UserPrincipalName,Name,DisplayName,SimpleDisplayName,PrimarySmtpAddress,Alias,SamAccountName,ExchangeGuid,RecipientTypeDetails,ForwardingAddress,ForwardingSmtpAddress,DeliverToMilboxAndForward,LitigationHoldEnabled,RetentionHoldEnabled,InPlaceHolds,GrantSendOnBehalfTo,HiddenFromAddressListsEnabled,ArchiveStatus,ArchiveName,ArchiveGuid,EmailAddresses,WhenChanged,WhenChangedUTC,WhenMailboxCreated,WhenCreated,WhenCreatedUTC,UMEnabled,ExternalOofOptions,IssueWarningQuota,ProhiitSendQuota,ProhibitSendReceiveQuota,UseDatabaseQuotaDefaults,MaxSendSize,MaxReceiveSize,CustomAttribute1,CustomAttribute2,CustomAttribute3,CustomAttribue4,CustomAttribute5,CustomAttribute6,CustomAttribute7,CustomAttribute8,CustomAttribute9,CustomAttribute10,CustomAttribute11,CustomAttribute12,CustomAttriute13,CustomAttribute14,CustomAttribute15,ExtensionCustomAttribute1,ExtensionCustomAttribute2,ExtensionCustomAttribute3,ExtensionCustomAttribute4,ExtensionCustomAttribute5)
}
$mailboxCount = $mailboxes.Count
$i = 1
ForEach ($mailbox in $mailboxes) {
    Write-Progress -Id 1 -Activity "EXO Mailbox Information Report" -Status "Processing $($i) of $($mailboxCount) Mailboxes --- $($mailbox.UserPrincipalName)" -PercentComplete (($i*100)/$mailboxCount)
    $mailboxInfo = Get-MailboxInformation $mailbox
    $output.Add([PSCustomObject]$mailboxInfo) | Out-Null
    $i++
}
$output | Select-Object UserPrincipalName,Name,DisplayName,SimpleDisplayName,PrimarySmtpAddress,Alias,SamAccountName,ExchangeGuid,RecipientTypeDetails,ForwardingAddress,ForwardingSmtpAddress,DeliverToMilboxAndForward,LitigationHoldEnabled,RetentionHoldEnabled,InPlaceHolds,GrantSendOnBehalfTo,HiddenFromAddressListsEnabled,ArchiveStatus,ArchiveName,ArchiveGuid,EmailAddresses,WhenChanged,WhenChangedUTC,WhenMailboxCreated,WhenCreated,WhenCreatedUTC,UMEnabled,ExternalOofOptions,IssueWarningQuota,ProhiitSendQuota,ProhibitSendReceiveQuota,UseDatabaseQuotaDefaults,MaxSendSize,MaxReceiveSize,CustomAttribute1,CustomAttribute2,CustomAttribute3,CustomAttribue4,CustomAttribute5,CustomAttribute6,CustomAttribute7,CustomAttribute8,CustomAttribute9,CustomAttribute10,CustomAttribute11,CustomAttribute12,CustomAttriute13,CustomAttribute14,CustomAttribute15,ExtensionCustomAttribute1,ExtensionCustomAttribute2,ExtensionCustomAttribute3,ExtensionCustomAttribute4,ExtensionCustomAttribute5 | Export-Csv $outputFile -NoClobber -NoTypeInformation -Encoding UTF8
