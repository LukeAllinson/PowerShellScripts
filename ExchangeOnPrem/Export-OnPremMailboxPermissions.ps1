<#
    .SYNOPSIS
        Name: Export-MailboxPermissions.ps1
        This script enumerates all permissions for every mailbox and exports to a csv file.

    .DESCRIPTION
        This script outputs permissions for each mailbox into a CSV

    .NOTES
        Authors: Luke Allinson (github:LukeAllinson)
                 Robin Dadswell (github:RobinDadswell)

    .PARAMETER OutputPath
        Full path to the folder where the output will be saved.
        Can be used without the parameter name in the first position only.

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

    .PARAMETER IncludeNoPermissions
        Includes mailboxes with no permissions in the export; by default only valid permissions are shown in the export.
        For example, if User01 has Full Access and SendOnBehalf permissions, then only these are shown in the report by default. If the IncludeNoPermissions parameter is included then SendAs permissions will also be included as "<none>".
        Similarly, if User02 has no permissions at all it will not be present in the export, however with this parameter set all three permissions will be included as "<none>".

    .EXAMPLE
        .\Export-MailboxPermissions.ps1 C:\Scripts\
        Exports mailbox permissions for all mailbox types

    .EXAMPLE
        .\Export-MailboxPermissions.ps1 -RecipientTypeDetails RoomMailbox,EquipmentMailbox -OutputPath C:\Scripts\ -IncludeNoPermissions
        Exports mailbox permissions only for Room and Equipment mailboxes; include all permissions, even if blank.

    .EXAMPLE
        .\Export-MailboxPermissions.ps1 C:\Scripts\ -MailboxFilter 'Department -eq "R&D"'
        Exports mailbox permissions for all mailboxes from the R&D department
#>

[CmdletBinding()]
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
    [switch]
    $IncludeNoPermissions,
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

function Resolve-Permissions
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory)]
        [System.Array]
        $Recipients,
        [Parameter(Mandatory)]
        [System.Array]
        $Groups,
        [System.Object]
        $Mailbox,
        [System.Object]
        $Permissions,
        [Parameter(Mandatory)]
        [string]
        [ValidateSet('FullAccess', 'SendAs', 'SendOnBehalf')]
        $PermissionType,
        [bool]
        $IncludeNoPermissions
    )

    $output = New-Object System.Collections.Generic.List[System.Object]
    if ($IncludeNoPermissions -and !$Permissions)
    {
        $permEntry = [ordered]@{
            UserPrincipalName           = $Mailbox.UserPrincipalName
            DisplayName                 = $Mailbox.DisplayName
            PrimarySmtpAddress          = $Mailbox.PrimarySmtpAddress
            SamAccountName              = $Mailbox.SamAccountName
            RecipientTypeDetails        = $Mailbox.RecipientTypeDetails
            PermissionType              = $PermissionType
            TrusteeIdentity             = '<NoDefinedPermissions>'
            TrusteeName                 = '<NoDefinedPermissions>'
            TrusteeSamAccountName       = '<NoDefinedPermissions>'
            TrusteeRecipientTypeDetails = '<NoDefinedPermissions>'
        }
        $output.Add([PSCustomObject]$permEntry) | Out-Null
        return $output
    }

    foreach ($perm in $Permissions)
    {
        switch ($PermissionType)
        {
            FullAccess
            {
                $permission = $null
                $permTrustee = $Recipients.Where({ ($_.Alias -eq $perm.User.ToString().Split('\')[1]) -or ($_.SamAccountName -eq $perm.User.ToString().Split('\')[1]) -or ($_.Name -eq $perm.User.ToString().Split('\')[1]) -or ($_.PrimarySmtpAddress -eq $perm.User.ToString().Split('\')[1]) -or ($_.emailaddresses -contains "smtp:$($perm.User.ToString().Split('\')[1])") })
                if (!$permTrustee)
                {
                    $permTrustee = $Groups.Where({ ($_.SamAccountName -eq $perm.User.ToString().Split('\')[1]) })
                    if ($permTrustee)
                    {
                        Switch ($permTrustee.GroupType)
                        {
                                ({ $PSItem -match 'BuiltinLocal' })
                            {
                                $permission = 'BuiltinLocal'
                            }
                                ({ $PSItem -match 'DomainLocal' })
                            {
                                $permission = 'DomainLocal'
                            }
                                ({ $PSItem -match 'Global' })
                            {
                                $permission = 'Global'
                            }
                                ({ $PSItem -match 'Universal' })
                            {
                                $permission = 'Universal'
                            }
                        }
                        if ($permTrustee.GroupType -match 'SecurityEnabled')
                        {
                            $permission = $permission + 'SecurityGroup'
                        }
                        else
                        {
                            $permission = $permission + 'DistributionGroup'
                        }
                    }
                }
                else
                {
                    $permission = $permTrustee.RecipientTypeDetails
                }
                $trusteeId = $perm.User
            }
            SendAs
            {
                $permission = $null
                $permTrustee = $Recipients.Where({ ($_.Alias -eq $perm.User.ToString().Split('\')[1]) -or ($_.SamAccountName -eq $perm.User.ToString().Split('\')[1]) -or ($_.Name -eq $perm.User.ToString().Split('\')[1]) -or ($_.PrimarySmtpAddress -eq $perm.User.ToString().Split('\')[1]) -or ($_.emailaddresses -contains "smtp:$($perm.User.ToString().Split('\')[1])") })
                if (!$permTrustee)
                {
                    $permTrustee = $Groups.Where({ ($_.SamAccountName -eq $perm.User.ToString().Split('\')[1]) })
                    if ($permTrustee)
                    {
                        Switch ($permTrustee.GroupType)
                        {
                                ({ $PSItem -match 'BuiltinLocal' })
                            {
                                $permission = 'BuiltinLocal'
                            }
                                ({ $PSItem -match 'DomainLocal' })
                            {
                                $permission = 'DomainLocal'
                            }
                                ({ $PSItem -match 'Global' })
                            {
                                $permission = 'Global'
                            }
                                ({ $PSItem -match 'Universal' })
                            {
                                $permission = 'Universal'
                            }
                        }
                        if ($permTrustee.GroupType -match 'SecurityEnabled')
                        {
                            $permission = $permission + 'SecurityGroup'
                        }
                        else
                        {
                            $permission = $permission + 'DistributionGroup'
                        }
                    }
                }
                else
                {
                    $permission = $permTrustee.RecipientTypeDetails
                }
                $trusteeId = $perm.User
            }
            SendOnBehalf
            {
                $permission = $null
                $permTrustee = $Recipients.Where({ $_.Name -eq $perm.Name })
                if (!$permTrustee)
                {
                    $permTrustee = $Groups.Where({ ($_.SamAccountName -eq $perm.Name) })
                    if ($permTrustee)
                    {
                        Switch ($permTrustee.GroupType)
                        {
                                ({ $PSItem -match 'BuiltinLocal' })
                            {
                                $permission = 'BuiltinLocal'
                            }
                                ({ $PSItem -match 'DomainLocal' })
                            {
                                $permission = 'DomainLocal'
                            }
                                ({ $PSItem -match 'Global' })
                            {
                                $permission = 'Global'
                            }
                                ({ $PSItem -match 'Universal' })
                            {
                                $permission = 'Universal'
                            }
                        }
                        if ($permTrustee.GroupType -match 'SecurityEnabled')
                        {
                            $permission = $permission + 'SecurityGroup'
                        }
                        else
                        {
                            $permission = $permission + 'DistributionGroup'
                        }
                    }
                }
                else
                {
                    $permission = $permTrustee.RecipientTypeDetails
                }
                $trusteeId = $perm
            }
        }
        if ($permTrustee)
        {
            $objPermEntry = [ordered]@{
                UserPrincipalName           = $Mailbox.UserPrincipalName
                DisplayName                 = $Mailbox.DisplayName
                PrimarySmtpAddress          = $Mailbox.PrimarySmtpAddress
                SamAccountName              = $Mailbox.SamAccountName
                RecipientTypeDetails        = $Mailbox.RecipientTypeDetails
                PermissionType              = $PermissionType
                TrusteeIdentity             = $permTrustee.PrimarySmtpAddress
                TrusteeName                 = $permTrustee.Name
                TrusteeSamAccountName       = $permTrustee.SamAccountName
                TrusteeRecipientTypeDetails = $permission
            }
            $output.Add([PSCustomObject]$objPermEntry) | Out-Null
        }
        else
        {
            $objPermEntry = [ordered]@{
                UserPrincipalName           = $Mailbox.UserPrincipalName
                DisplayName                 = $Mailbox.DisplayName
                PrimarySmtpAddress          = $Mailbox.PrimarySmtpAddress
                SamAccountName              = $Mailbox.SamAccountName
                RecipientTypeDetails        = $Mailbox.RecipientTypeDetails
                PermissionType              = $PermissionType
                TrusteeIdentity             = $trusteeId
                TrusteeName                 = '<TrusteeNotFound>'
                TrusteeSamAccountName       = '<TrusteeNotFound>'
                TrusteeRecipientTypeDetails = '<TrusteeNotFound>'
            }
            $output.Add([PSCustomObject]$objPermEntry) | Out-Null
        }
    }
    return $output
}

### Main Script
# Set Constants and Variables
if (!$PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent)
{
    Write-Progress -Id 1 -Activity 'Mailbox Permissions Report' -Status 'Initialising...' -PercentComplete 15
}
$i = 1
$timeStamp = Get-Date -Format yyyyMMdd-HHmm
$outputFile = $OutputPath.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar + 'MailboxPermissions_' + $timeStamp + '.csv'
$output = New-Object System.Collections.Generic.List[System.Object]

# Get In-Scope Mailboxes
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

if (!$PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent)
{
    Write-Progress -Id 1 -Activity 'Mailbox Permissions Report' -Status 'Getting In-Scope Mailboxes' -PercentComplete 30
}

try
{
    Write-Verbose 'Getting in-scope Mailboxes from Exchange'
    $mailboxes = @(Get-Mailbox @commandHashTable | Sort-Object UserPrincipalName)
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

# Get all Recipients
$commandHashTable2 = @{
    ResultSize  = 'Unlimited'
    ErrorAction = 'Stop'
}

if (!$PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent)
{
    Write-Progress -Id 1 -Activity 'Mailbox Permissions Report' -Status 'Getting All Recipients' -PercentComplete 45
}

try
{
    Write-Verbose 'Getting all Recipients from Exchange'
    $recipients = @(Get-Recipient @commandHashTable2 | Sort-Object Name)
}
catch
{
    throw
}

if (!$PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent)
{
    Write-Progress -Id 1 -Activity 'Mailbox Permissions Report' -Status 'Getting All Groups' -PercentComplete 55
}

try
{
    Write-Verbose 'Getting all Recipients from Exchange'
    $groups = @(Get-Group @commandHashTable2 | Sort-Object Name)
}
catch
{
    throw
}


#  Loop through the list of mailboxes and output the results to the CSV
if (!$PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent)
{
    Write-Progress -Id 1 -Activity 'Mailbox Permissions Report' -Status 'Processing Mailbox Permissions' -PercentComplete 75
}

foreach ($mailbox in $mailboxes)
{
    if (!$PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent)
    {
        Write-Progress -Id 2 -ParentId 1 -Activity 'Processing Mailbox Permissions' -Status "Processing $($i) of $($mailboxCount) Mailboxes --- $($mailbox.UserPrincipalName)" -PercentComplete (($i * 100) / $mailboxCount)
        $i++
    }

    # Full Access Permissions
    Write-Verbose "Processing FullAccess permissions for $($mailbox.UserPrincipalName)"
    try
    {
        $fullAccessPerms = @(Get-MailboxPermission $mailbox.Identity -ErrorAction stop).Where({ ($_.AccessRights -like 'Full*') -and ($_.User -notmatch 'SELF') -and ($_.IsInherited -eq $false) })
    }
    catch
    {
        Write-Verbose "Failure getting FullAccess permissions for $($mailbox.UserPrincipalName)"
        $faPermEntry = [ordered]@{
            UserPrincipalName     = $mailbox.UserPrincipalName
            DisplayName           = $mailbox.DisplayName
            PrimarySmtpAddress    = $mailbox.PrimarySmtpAddress
            SamAccountName        = $mailbox.SamAccountName
            RecipientTypeDetails  = $mailbox.RecipientTypeDetails
            PermissionType        = 'FullAccess'
            TrusteeUPN            = '<ErrorRunningCommand>'
            TrusteeDisplayName    = '<ErrorRunningCommand>'
            TrusteeSamAccountName = '<ErrorRunningCommand>'
            TrusteeStatus         = '<ErrorRunningCommand>'
        }
        $output.Add([PSCustomObject]$faPermEntry) | Out-Null
        Continue
    }

    $resolvedFullAccessPerms = [Object[]](Resolve-Permissions -Recipients $recipients -Groups $groups -Mailbox $mailbox -Permissions $fullAccessPerms -PermissionType 'FullAccess' -IncludeNoPermissions $IncludeNoPermissions)
    if ($resolvedFullAccessPerms)
    {
        $output.AddRange($resolvedFullAccessPerms)
    }

    # SendAs Permissions
    Write-Verbose "Processing SendAs permissions for $($mailbox.UserPrincipalName)"
    try
    {
        $sendAsPerms = @(Get-ADPermission -Identity $mailbox.Name).Where({ ($_.ExtendedRights -like '*send*') -and ($_.User -notmatch 'SELF') })
    }
    catch
    {
        Write-Verbose "Failure getting SendAs permissions for $($mailbox.UserPrincipalName)"
        $saPermEntry = [ordered]@{
            UserPrincipalName     = $mailbox.UserPrincipalName
            DisplayName           = $mailbox.DisplayName
            PrimarySmtpAddress    = $mailbox.PrimarySmtpAddress
            SamAccountName        = $mailbox.SamAccountName
            RecipientTypeDetails  = $mailbox.RecipientTypeDetails
            PermissionType        = 'SendAs'
            TrusteeUPN            = '<ErrorRunningCommand>'
            TrusteeDisplayName    = '<ErrorRunningCommand>'
            TrusteeSamAccountName = '<ErrorRunningCommand>'
            TrusteeStatus         = '<ErrorRunningCommand>'
        }
        $output.Add([PSCustomObject]$saPermEntry) | Out-Null
        Continue
    }
    $resolvedSendAsPerms = [Object[]](Resolve-Permissions -Recipients $recipients -Groups $groups -Mailbox $mailbox -Permissions $sendAsPerms -PermissionType 'SendAs' -IncludeNoPermissions $IncludeNoPermissions)
    if ($resolvedSendAsPerms)
    {
        $output.AddRange($resolvedSendAsPerms)
    }

    # SendOnBehalf Permissions
    Write-Verbose "Processing SendOnBehalfTo permissions for $($mailbox.UserPrincipalName)"
    $sendOnBehalfPerms = $mailbox.GrantSendOnBehalfTo
    $resolvedSendOnBehalfPerms = [Object[]](Resolve-Permissions -Recipients $recipients -Groups $groups -Mailbox $mailbox -Permissions $sendOnBehalfPerms -PermissionType 'SendOnBehalf' -IncludeNoPermissions $IncludeNoPermissions)
    if ($resolvedSendOnBehalfPerms)
    {
        $output.AddRange($resolvedSendOnBehalfPerms)
    }
}

if (!$PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent)
{
    Write-Progress -Id 2 -ParentId 1 -Activity 'Processing Mailbox Permissions' -Completed
    Write-Progress -Id 1 -Activity 'Mailbox Permissions Report' -Status 'Writing Output' -PercentComplete 95
}

# Export to csv
$output | Export-Csv $outputFile -NoClobber -NoTypeInformation -Encoding UTF8

if (!$PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent)
{
    Write-Progress -Id 1 -Activity 'Mailbox Permissions Report' -Completed
}

return "Mailbox permissions have been exported to $outputfile"
