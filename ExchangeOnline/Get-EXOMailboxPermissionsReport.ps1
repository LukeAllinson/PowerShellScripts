#Requires -Version 5 -Modules ExchangeOnlineManagement

<#
	.SYNOPSIS
		Name: Get-EXOMailboxPermissionsReport.ps1
		This script enumerates all permissions for every mailbox into a csv

	.DESCRIPTION
		This script connects to EXO and then outputs permissions for each mailbox into a CSV 

	.NOTES
		Version: 0.2
        Updated: 14-10-2021 v0.2    Updated to use Rest-based commands where possible
		Updated: 01-05-2021	v0.1	Initial draft

		Authors: Luke Allinson, Robin Dadswell
#>

Function _Progress {
    Param ($PercentComplete,$Status)
    Write-Progress -Id 1 -Activity "EXO Shared Mailbox Permissions Report" -Status $Status -PercentComplete ($PercentComplete)
} #End Function _ParentProgress

# Check if there is an active Exchange Online PowerShell session
$PSSessions = Get-PSSession | Select-Object -Property State, Name
If (((@($PSSessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -gt 0) -ne $true) {
    Write-Host "Exchange Online PowerShell not connected..." -ForegroundColor Yellow
    Write-Host "Connecting..." -ForegroundColor Yellow
    Connect-ExchangeOnline
} Else {
    Write-Host "Already connected to Exchange Online PowerShell" -ForegroundColor Green
}

# Print screen output description
Write-Host
Write-Host "Mailbox Permissions Report"
Write-Host "---------------------------------"
Write-Host "Mailbox being processed in Green" -ForegroundColor Green
Write-Host "Full Mailbox Permissions in Cyan" -ForegroundColor Cyan
Write-Host "SendAs Permissions in Magenta" -ForegroundColor Magenta
Write-Host "SendOnBehalf Permissions in Yellow" -ForegroundColor Yellow
Start-Sleep -Seconds 5

# Set Constants and Variables
$i = 1
$Date = Get-Date -Format ddMMyyyy-HHmm
$Output = [System.Collections.ArrayList]@()
Write-Host
Write-Host "Getting all Mailboxes..."
Write-Host
$Mailboxes = Get-EXOMailbox -Resultsize Unlimited -Properties IsDirSynced,GrantSendOnBehalfTo | Sort-Object UserPrincipalName
$MailboxCount = $Mailboxes.Count
$AllSendAsPerms = Get-EXORecipientPermission -ResultSize unlimited | Where-Object {$_.AccessRights -eq "SendAs" -and $_.Trustee -notmatch "SELF"}
$i = 1
ForEach ($MB in $Mailboxes) {
    _Progress (($i*100)/$MailboxCount) "Processing $($i) of $($MailboxCount) Mailboxes --- $($MB.UserPrincipalName)"
    If (!($MB.IsInactiveMailbox)) {
        Write-Host "$($i) of $($MailboxCount) --- $($MB.UserPrincipalName)" -ForegroundColor Green
        # Get Full Access Permissions
        $FullAccessPerms = Get-EXOMailboxPermission $MB.Identity | Where-Object {($_.AccessRights -like 'Full*') -and ($_.User -notlike "*SELF*")}
        If ($FullAccessPerms) {
            ForEach ($FAPerm in $FullAccessPerms) {
                Write-Host "--- " $FAPerm.User -ForegroundColor Cyan
                $FAUser = $Mailboxes | Where-Object {$_.UserPrincipalName -eq $FAPerm.User}
                If ($FAUser) {
                    If ($FAUser.IsInactiveMailbox -eq $TRUE) {
                        $FAPermObj = [System.Collections.ArrayList]@()
                        $FAPermObj = @{
                            UserPrincipalName = $MB.UserPrincipalName
                            DisplayName = $MB.DisplayName
                            PrimarySmtpAddress = $MB.PrimarySmtpAddress
                            RecipientTypeDetails = $MB.RecipientTypeDetails
                            IsDirSynced = $MB.IsDirSynced
                            PermissionType = "FullAccess"
                            TrusteeUPN = $FAUser.UserPrincipalName
                            TrusteeDisplayName = $FAUser.DisplayName
                            TrusteeStatus = "Inactive"
                        }
                        $Output.Add([PSCustomObject]$FAPermObj) | Out-Null
                    } Else {
                        $FAPermObj = [System.Collections.ArrayList]@()
                        $FAPermObj = @{
                            UserPrincipalName = $MB.UserPrincipalName
                            DisplayName = $MB.DisplayName
                            PrimarySmtpAddress = $MB.PrimarySmtpAddress
                            RecipientTypeDetails = $MB.RecipientTypeDetails
                            IsDirSynced = $MB.IsDirSynced
                            PermissionType = "FullAccess"
                            TrusteeUPN = $FAUser.UserPrincipalName
                            TrusteeDisplayName = $FAUser.DisplayName
                            TrusteeStatus = "Active"
                        }
                        $Output.Add([PSCustomObject]$FAPermObj) | Out-Null
                    }
                }
            }
        }

        # Get SendAs Permissions
        $SendAsPerms = $AllSendAsPerms | Where-Object {$_.Identity -eq $MB.Identity}
        If ($SendAsPerms) {
            ForEach ($SAPerm in $SendAsPerms) {
                Write-Host "--- " $SAPerm.Trustee -ForegroundColor Magenta
                $SAUser = $Mailboxes | Where-Object {$_.UserPrincipalName -eq $SAPerm.Trustee}
                If ($SAUser) {
                    If ($SAUser.IsInactiveMailbox -eq $TRUE) {
                        $SAPermObj = [System.Collections.ArrayList]@()
                        $SAPermObj = @{
                            UserPrincipalName = $MB.UserPrincipalName
                            DisplayName = $MB.DisplayName
                            PrimarySmtpAddress = $MB.PrimarySmtpAddress
                            RecipientTypeDetails = $MB.RecipientTypeDetails
                            IsDirSynced = $MB.IsDirSynced
                            PermissionType = "SendAs"
                            TrusteeUPN = $SAUser.UserPrincipalName
                            TrusteeDisplayName = $SAUser.DisplayName
                            TrusteeStatus = "Inactive"
                        }
                        $Output.Add([PSCustomObject]$FAPermObj) | Out-Null
                    } Else {
                        $SAPermObj = [System.Collections.ArrayList]@()
                        $SAPermObj = @{
                            UserPrincipalName = $MB.UserPrincipalName
                            DisplayName = $MB.DisplayName
                            PrimarySmtpAddress = $MB.PrimarySmtpAddress
                            RecipientTypeDetails = $MB.RecipientTypeDetails
                            IsDirSynced = $MB.IsDirSynced
                            PermissionType = "SendAs"
                            TrusteeUPN = $SAUser.UserPrincipalName
                            TrusteeDisplayName = $SAUser.DisplayName
                            TrusteeStatus = "Active"
                        }
                        $Output.Add([PSCustomObject]$SAPermObj) | Out-Null
                    }
                }
            }
        }
        # SendOnBehalf Permissions
        $SendOnBehalfPerms = $MB.GrantSendOnBehalfTo
        If ($SendOnBehalfPerms) {
            ForEach ($SOBPerm in $SendOnBehalfPerms) {
                Write-Host "--- " $SOBPerm -ForegroundColor Yellow
                $SOBUser = $Mailboxes | Where-Object {$_.Name -eq $SOBPerm}
                If ($SOBUser) {
                    If ($SOBUser.IsInactiveMailbox -eq $TRUE) {
                        $SOBPermObj = [System.Collections.ArrayList]@()
                        $SOBPermObj = @{
                            UserPrincipalName = $MB.UserPrincipalName
                            DisplayName = $MB.DisplayName
                            PrimarySmtpAddress = $MB.PrimarySmtpAddress
                            RecipientTypeDetails = $MB.RecipientTypeDetails
                            IsDirSynced = $MB.IsDirSynced
                            PermissionType = "SendOnBehalf"
                            TrusteeUPN = $SOBUser.UserPrincipalName
                            TrusteeDisplayName = $SOBUser.DisplayName
                            TrusteeStatus = "Inactive"
                        }
                        $Output.Add([PSCustomObject]$FAPermObj) | Out-Null
                    } Else {
                        $SOBPermObj = [System.Collections.ArrayList]@()
                        $SOBPermObj = @{
                            UserPrincipalName = $MB.UserPrincipalName
                            DisplayName = $MB.DisplayName
                            PrimarySmtpAddress = $MB.PrimarySmtpAddress
                            RecipientTypeDetails = $MB.RecipientTypeDetails
                            IsDirSynced = $MB.IsDirSynced
                            PermissionType = "SendOnBehalf"
                            TrusteeUPN = $SOBUser.UserPrincipalName
                            TrusteeDisplayName = $SOBUser.DisplayName
                            TrusteeStatus = "Active"
                        }
                        $Output.Add([PSCustomObject]$SOBPermObj) | Out-Null
                    }
                }

            }
        }
        If (!($Output.UserPrincipalName -contains $MB.UserPrincipalName)) {
            Write-Host "--- No permissions found"
            $SOBPermObj = [System.Collections.ArrayList]@()
            $SOBPermObj = @{
                UserPrincipalName = $MB.UserPrincipalName
                DisplayName = $MB.DisplayName
                PrimarySmtpAddress = $MB.PrimarySmtpAddress
                RecipientTypeDetails = $MB.RecipientTypeDetails
                IsDirSynced = $MB.IsDirSynced
                PermissionType = "<no permissions found>"
                TrusteeUPN = $NULL
                TrusteeDisplayName = $NULL
                TrusteeStatus = $NULL
            }
            $Output.Add([PSCustomObject]$SOBPermObj) | Out-Null
        }
        $i++
    }
}
$Output | Select-Object UserPrincipalName,DisplayName,PrimarySmtpAddress,RecipientTypeDetails,IsDirSynced,PermissionType,TrusteeUPN,TrusteeDisplayName | Export-Csv .\ExOMailbox_PermissionsReport_$Date.csv -NoClobber -NoTypeInformation -Encoding UTF8
Write-Host "CSV File saved: ExOMailbox_PermissionsReport_$Date.csv"
