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

# Check if there is an active Exchange Online PowerShell session
$PSSessions = Get-PSSession | Select-Object -Property State, Name
if (((@($PSSessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -gt 0) -ne $true) {
    Write-Host "Exchange Online PowerShell not connected..." -ForegroundColor Yellow
    Write-Host "Connecting..." -ForegroundColor Yellow
    Connect-ExchangeOnline
} else {
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
$timeStamp = Get-Date -Format ddMMyyyy-HHmm
$tenantName = (Get-OrganizationConfig).Name.Split(".")[0]
$outputFile = $OutputPath.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar + $timeStamp + '-' + $tenantName + '-' + 'EXOMailboxInformation.csv'
$output = [System.Collections.ArrayList]@()
Write-Host
Write-Host "Getting all Mailboxes..."
Write-Host
$mailboxes = Get-EXOMailbox -Resultsize Unlimited -Properties IsDirSynced,GrantSendOnBehalfTo | Sort-Object UserPrincipalName
$mailboxCount = $mailboxes.Count
$allSendAsPerms = Get-EXORecipientPermission -ResultSize unlimited | Where-Object {$_.AccessRights -eq "SendAs" -and $_.Trustee -notmatch "SELF"}
$i = 1
foreach ($mailbox in $mailboxes) {
    Write-Progress -Id 1 -Activity "EXO Mailbox Permissions Report" -Status "Processing $($i) of $($mailboxCount) Mailboxes --- $($mailbox.UserPrincipalName)" -PercentComplete (($i*100)/$mailboxCount)
    if (!($mailbox.IsInactiveMailbox)) {
        Write-Host "$($i) of $($mailboxCount) --- $($mailbox.UserPrincipalName)" -ForegroundColor Green
        # Get Full Access Permissions
        $fullAccessPerms = Get-EXOMailboxPermission $mailbox.Identity | Where-Object {($_.AccessRights -like 'Full*') -and ($_.User -notlike "*SELF*")}
        if ($fullAccessPerms) {
            foreach ($faPerm in $fullAccessPerms) {
                Write-Host "--- " $faPerm.User -ForegroundColor Cyan
                $faUser = $mailboxes | Where-Object {$_.UserPrincipalName -eq $faPerm.User}
                if ($faUser) {
                    if ($faUser.IsInactiveMailbox -eq $true) {
                        $faPermEntry = @{
                            UserPrincipalName = $mailbox.UserPrincipalName
                            DisplayName = $mailbox.DisplayName
                            PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                            RecipientTypeDetails = $mailbox.RecipientTypeDetails
                            IsDirSynced = $mailbox.IsDirSynced
                            PermissionType = "FullAccess"
                            TrusteeUPN = $faUser.UserPrincipalName
                            TrusteeDisplayName = $faUser.DisplayName
                            TrusteeStatus = "Inactive"
                        }
                        $output.Add([PSCustomObject]$faPermEntry) | Out-Null
                    } else {
                        $faPermEntry = @{
                            UserPrincipalName = $mailbox.UserPrincipalName
                            DisplayName = $mailbox.DisplayName
                            PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                            RecipientTypeDetails = $mailbox.RecipientTypeDetails
                            IsDirSynced = $mailbox.IsDirSynced
                            PermissionType = "FullAccess"
                            TrusteeUPN = $faUser.UserPrincipalName
                            TrusteeDisplayName = $faUser.DisplayName
                            TrusteeStatus = "Active"
                        }
                        $output.Add([PSCustomObject]$faPermEntry) | Out-Null
                    }
                }
            }
        }

        # Get SendAs Permissions
        $sendAsPerms = $allSendAsPerms | Where-Object {$_.Identity -eq $mailbox.Identity}
        if ($sendAsPerms) {
            foreach ($saPerm in $sendAsPerms) {
                Write-Host "--- " $saPerm.Trustee -ForegroundColor Magenta
                $saUser = $mailboxes | Where-Object {$_.UserPrincipalName -eq $saPerm.Trustee}
                if ($saUser) {
                    if ($saUser.IsInactiveMailbox -eq $TRUE) {
                        $saPermEntry = @{
                            UserPrincipalName = $mailbox.UserPrincipalName
                            DisplayName = $mailbox.DisplayName
                            PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                            RecipientTypeDetails = $mailbox.RecipientTypeDetails
                            IsDirSynced = $mailbox.IsDirSynced
                            PermissionType = "SendAs"
                            TrusteeUPN = $saUser.UserPrincipalName
                            TrusteeDisplayName = $saUser.DisplayName
                            TrusteeStatus = "Inactive"
                        }
                        $output.Add([PSCustomObject]$saPermEntry) | Out-Null
                    } else {
                        $saPermEntry = @{
                            UserPrincipalName = $mailbox.UserPrincipalName
                            DisplayName = $mailbox.DisplayName
                            PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                            RecipientTypeDetails = $mailbox.RecipientTypeDetails
                            IsDirSynced = $mailbox.IsDirSynced
                            PermissionType = "SendAs"
                            TrusteeUPN = $saUser.UserPrincipalName
                            TrusteeDisplayName = $saUser.DisplayName
                            TrusteeStatus = "Active"
                        }
                        $output.Add([PSCustomObject]$saPermEntry) | Out-Null
                    }
                }
            }
        }
        # SendOnBehalf Permissions
        $sendOnBehalfPerms = $mailbox.GrantSendOnBehalfTo
        if ($sendOnBehalfPerms) {
            foreach ($sobPerms in $sendOnBehalfPerms) {
                Write-Host "--- " $sobPerms -ForegroundColor Yellow
                $sobUsers = $mailboxes | Where-Object {$_.Name -eq $sobPerms}
                if ($sobUsers) {
                    if ($sobUsers.IsInactiveMailbox -eq $TRUE) {
                        $sobPermEntry = @{
                            UserPrincipalName = $mailbox.UserPrincipalName
                            DisplayName = $mailbox.DisplayName
                            PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                            RecipientTypeDetails = $mailbox.RecipientTypeDetails
                            IsDirSynced = $mailbox.IsDirSynced
                            PermissionType = "SendOnBehalf"
                            TrusteeUPN = $sobUsers.UserPrincipalName
                            TrusteeDisplayName = $sobUsers.DisplayName
                            TrusteeStatus = "Inactive"
                        }
                        $output.Add([PSCustomObject]$sobPermEntry) | Out-Null
                    } else {
                        $sobPermEntry = @{
                            UserPrincipalName = $mailbox.UserPrincipalName
                            DisplayName = $mailbox.DisplayName
                            PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                            RecipientTypeDetails = $mailbox.RecipientTypeDetails
                            IsDirSynced = $mailbox.IsDirSynced
                            PermissionType = "SendOnBehalf"
                            TrusteeUPN = $sobUsers.UserPrincipalName
                            TrusteeDisplayName = $sobUsers.DisplayName
                            TrusteeStatus = "Active"
                        }
                        $output.Add([PSCustomObject]$sobPermEntry) | Out-Null
                    }
                }

            }
        }
        # No Permissions found
        if (!($output.UserPrincipalName -contains $mailbox.UserPrincipalName)) {
            Write-Host "--- No permissions found"
            $noPermEntry = @{
                UserPrincipalName = $mailbox.UserPrincipalName
                DisplayName = $mailbox.DisplayName
                PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                RecipientTypeDetails = $mailbox.RecipientTypeDetails
                IsDirSynced = $mailbox.IsDirSynced
                PermissionType = "<no permissions found>"
                TrusteeUPN = $NULL
                TrusteeDisplayName = $NULL
                TrusteeStatus = $NULL
            }
            $output.Add([PSCustomObject]$noPermEntry) | Out-Null
        }
        $i++
    }
}
$output | Select-Object UserPrincipalName,DisplayName,PrimarySmtpAddress,RecipientTypeDetails,IsDirSynced,PermissionType,TrusteeUPN,TrusteeDisplayName | Export-Csv $outputFile -NoClobber -NoTypeInformation -Encoding UTF8
Write-Host "CSV File saved: $outputFile"
