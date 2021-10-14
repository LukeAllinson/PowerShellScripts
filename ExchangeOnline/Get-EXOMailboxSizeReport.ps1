Function _Progress {
    Param ($PercentComplete,$Status)
    Write-Progress -Id 1 -Activity "Exchange EXO Shared Mailbox Size Report" -Status $Status -PercentComplete ($PercentComplete)
} #End Function _ParentProgress
Function Get-MailboxInformation ($Mailbox) {
    # Get Mailbox Statistics
    $PrimaryStats = Get-MailboxStatistics -Identity $Mailbox.DistinguishedName -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
    $PrimaryTotalItemSizeMB = $PrimaryStats | Select-Object @{name=”TotalItemSizeMB”; expression={[math]::Round(($_.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1MB),2)}}
    $ArchiveStats = Get-MailboxStatistics -Identity $Mailbox.DistinguishedName -Archive -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
    $ArchiveTotalItemSizeMB = $ArchiveStats | Select-Object @{name=”TotalItemSizeMB”; expression={[math]::Round(($_.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1MB),2)}}
    # Store everything in an Arraylist
    $UserObj = [System.Collections.ArrayList]@()
    $UserObj = @{
        UPN = $Mailbox.UserPrincipalName
        Name = $Mailbox.Name
        DisplayName = $Mailbox.Displayname
        SimpleDisplayName = $Mailbox.SimpleDisplayName
        PrimarySmtpAddress = $Mailbox.PrimarySmtpAddress
        Alias = $Mailbox.Alias
        SamAccountName = $Mailbox.SamAccountName
        RecipientTypeDetails = $Mailbox.RecipientTypeDetails
        ForwardingAddress = $Mailbox.ForwardingAddress
        ForwardingSmtpAddress = $Mailbox.ForwardingSmtpAddress
        DeliverToMailboxAndForward = $Mailbox.DeliverToMailboxAndForward
        LitigationHoldEnabled = $Mailbox.LitigationHoldEnabled
        RetentionHoldEnabled = $Mailbox.RetentionHoldEnabled
        InPlaceHolds = $Mailbox.InPlaceHolds
        GrantSendOnBehalfTo = $Mailbox.GrantSendOnBehalfTo
        HiddenFromAddressListsEnabled = $Mailbox.HiddenFromAddressListsEnabled
        ExchangeGuid = $Mailbox.ExchangeGuid
        ArchiveStatus = $Mailbox.ArchiveStatus
        ArchiveName = $Mailbox.ArchiveName
        ArchiveGuid = $Mailbox.ArchiveGuid
        EmailAddresses = ($Mailbox.EmailAddresses -join ";")
        WhenChanged = $Mailbox.WhenChanged
        WhenChangedUTC = $Mailbox.WhenChangedUTC
        WhenMailboxCreated = $Mailbox.WhenMailboxCreated
        WhenCreated = $Mailbox.WhenCreated
        WhenCreatedUTC = $Mailbox.WhenCreatedUTC
        UMEnabled = $Mailbox.UMEnabled
        ExternalOofOptions = $Mailbox.ExternalOofOptions
        IssueWarningQuota = $Mailbox.IssueWarningQuota
        ProhibitSendQuota = $Mailbox.ProhibitSendQuota
        ProhibitSendReceiveQuota = $Mailbox.ProhibitSendReceiveQuota
        UseDatabaseQuotaDefaults = $Mailbox.UseDatabaseQuotaDefaults
        MaxSendSize = $Mailbox.MaxSendSize
        MaxReceiveSize = $Mailbox.MaxReceiveSize
        CustomAttribute1 = $Mailbox.CustomAttribute1
        CustomAttribute2 = $Mailbox.CustomAttribute2
        CustomAttribute3 = $Mailbox.CustomAttribute3
        CustomAttribute4 = $Mailbox.CustomAttribute4
        CustomAttribute5 = $Mailbox.CustomAttribute5
        CustomAttribute6 = $Mailbox.CustomAttribute6
        CustomAttribute7 = $Mailbox.CustomAttribute7
        CustomAttribute8 = $Mailbox.CustomAttribute8
        CustomAttribute9 = $Mailbox.CustomAttribute9
        CustomAttribute10 = $Mailbox.CustomAttribute10
        CustomAttribute11 = $Mailbox.CustomAttribute11
        CustomAttribute12 = $Mailbox.CustomAttribute12
        CustomAttribute13 = $Mailbox.CustomAttribute13
        CustomAttribute14 = $Mailbox.CustomAttribute14
        CustomAttribute15 = $Mailbox.CustomAttribute15
        ExtensionCustomAttribute1 = $Mailbox.ExtensionCustomAttribute1
        ExtensionCustomAttribute2 = $Mailbox.ExtensionCustomAttribute2
        ExtensionCustomAttribute3 = $Mailbox.ExtensionCustomAttribute3
        ExtensionCustomAttribute4 = $Mailbox.ExtensionCustomAttribute4
        ExtensionCustomAttribute5 = $Mailbox.ExtensionCustomAttribute5
    }
    If ($PrimaryStats) {
        $UserObj["TotalItemSize(MB)"] = $PrimaryTotalItemSizeMB.TotalItemSizeMB
        $UserObj["ItemCount"] = $PrimaryStats.ItemCount
        $UserObj["DeletedItemCount"] = $PrimaryStats.DeletedItemCount
        $UserObj["LastLogonTime"] = $PrimaryStats.LastLogonTime
    }
    If ($ArchiveStats) {
        $UserObj["Archive_TotalItemSize(MB)"] = $ArchiveTotalItemSizeMB.TotalItemSizeMB
        $UserObj["Archive_ItemCount"] = $ArchiveStats.ItemCount
        $UserObj["Archive_DeletedItemCount"] = $ArchiveStats.DeletedItemCount
        $UserObj["Archive_LastLogonTime"] = $ArchiveStats.LastLogonTime
    }
    Return $UserObj
} #End Function Get-MailboxInformation

# Main Script
$i = 1
$Date = Get-Date -Format ddMMyyyy-HHmm
$Output = [System.Collections.ArrayList]@()
$Mailboxes = @(Get-Mailbox -Resultsize Unlimited)
$MailboxCount = $Mailboxes.Count
$i = 1
ForEach ($MB in $Mailboxes) {
    _Progress (($i*100)/$MailboxCount) "Processing $($i) of $($MailboxCount) Mailboxes --- $($MB.UserPrincipalName)"
    $MailboxInfo = Get-MailboxInformation $MB
    $Output.Add([PSCustomObject]$MailboxInfo) | Out-Null
    $i++
}
$Output | Select-Object UPN,Name,DisplayName,SimpleDisplayName,PrimarySmtpAddress,Alias,SamAccountName,RecipientTypeDetails,ForwardingAddress,ForwardingSmtpAddress,DeliverToMilboxAndForward,LitigationHoldEnabled,RetentionHoldEnabled,InPlaceHolds,GrantSendOnBehalfTo,HiddenFromAddressListsEnabled,ExchangeGuid,"TotalItemSize(MB)",ItemCount,DeletedItemCount,LastLogonTime,ArchiveStatus,ArchiveName,ArchiveGuid,"Archive_TotalItemSize(MB)",Archive_ItemCount,Archive_DeletedItemCount,Arhive_LastLogonTime,EmailAddresses,WhenChanged,WhenChangedUTC,WhenMailboxCreated,WhenCreated,WhenCreatedUTC,UMEnabled,ExternalOofOptions,IssueWarningQuota,ProhiitSendQuota,ProhibitSendReceiveQuota,UseDatabaseQuotaDefaults,MaxSendSize,MaxReceiveSize,CustomAttribute1,CustomAttribute2,CustomAttribute3,CustomAttribue4,CustomAttribute5,CustomAttribute6,CustomAttribute7,CustomAttribute8,CustomAttribute9,CustomAttribute10,CustomAttribute11,CustomAttribute12,CustomAttriute13,CustomAttribute14,CustomAttribute15,ExtensionCustomAttribute1,ExtensionCustomAttribute2,ExtensionCustomAttribute3,ExtensionCustomAttribute4,ExtensionCustomAttribute5 | Export-Csv .\ExO_MailboxSizeReport_$Date.csv -NoClobber -NoTypeInformation -Encoding UTF8
