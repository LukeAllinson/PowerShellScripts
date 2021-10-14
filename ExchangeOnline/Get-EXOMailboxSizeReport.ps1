#Requires -Version 5 -Modules ExchangeOnlineManagement

<#
	.SYNOPSIS
		Name: Get-EXOMailboxSizeReport.ps1
		This gathers mailbox information including primary and archive size and item count.

	.DESCRIPTION
		This script connects to EXO and then outputs Mailbox information and statistics to a CSV file. 

	.NOTES
		Version: 0.2
        Updated: 14-10-2021 v0.2    Rewritten to improve speed, remove superflous information
		Updated: <unknown>	v0.1	Initial draft

		Authors: Luke Allinson, Robin Dadswell
#>

param
(
    [Parameter(
        Mandatory
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
    $OutputPath
)

function Get-MailboxInformation ($mailbox) {
    # Get Mailbox Statistics
    $primaryStats = Get-EXOMailboxStatistics -Identity $mailbox.Guid -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
    $primaryTotalItemSizeMB = $primaryStats | Select-Object @{name=”TotalItemSizeMB”; expression={[math]::Round(($_.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1MB),2)}}
    # If an Archive exists, then get Statistics
    if ($mailbox.ArchiveStatus -ne "None") {
        $archiveStats = Get-EXOMailboxStatistics -Identity $mailbox.Guid -Archive -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
        $archiveTotalItemSizeMB = $archiveStats | Select-Object @{name=”TotalItemSizeMB”; expression={[math]::Round(($_.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1MB),2)}}
    }
    # Store everything in an Arraylist
    $mailboxInfo = @{
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
    return [PSCustomObject]$mailboxInfo
} #End Function Get-MailboxInformation

# Main Script
$timeStamp = Get-Date -Format ddMMyyyy-HHmm
$tenantName = (Get-OrganizationConfig).Name.Split(".")[0]
$outputFile = $OutputPath.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar + $timeStamp + '-' + $tenantName + '-' + 'EXOMailboxSizeReport.csv'
$output = [System.Collections.ArrayList]@()
$mailboxes = @(Get-EXOMailbox -Resultsize Unlimited -Properties LitigationHoldEnabled,RetentionHoldEnabled,InPlaceHolds,ArchiveStatus)
$mailboxCount = $mailboxes.Count
$i = 1
foreach ($mailbox in $mailboxes) {
    Write-Progress -Id 1 -Activity "EXO Mailbox Size Report" -Status "Processing $($i) of $($mailboxCount) Mailboxes --- $($mailbox.UserPrincipalName)" -PercentComplete (($i*100)/$mailboxCount)
    $mailboxInfo = Get-MailboxInformation $mailbox
    $output.Add([PSCustomObject]$mailboxInfo) | Out-Null
    $i++
}
$output | Select-Object UserPrincipalName,DisplayName,PrimarySmtpAddress,Alias,RecipientTypeDetails,LitigationHoldEnabled,RetentionHoldEnabled,InPlaceHolds,"TotalItemSize(MB)",ItemCount,DeletedItemCount,LastLogonTime,ArchiveStatus,"Archive_TotalItemSize(MB)",Archive_ItemCount,Archive_DeletedItemCount,Arhive_LastLogonTime | Export-Csv $outputFile -NoClobber -NoTypeInformation -Encoding UTF8
