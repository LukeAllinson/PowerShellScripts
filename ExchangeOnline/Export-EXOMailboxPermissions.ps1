#Requires -Version 5 -Modules ExchangeOnlineManagement

<#
    .SYNOPSIS
        Name: Export-EXOMailboxPermissions.ps1
        This script enumerates all permissions for every mailbox and exports to a csv file.

    .DESCRIPTION
        This script connects to EXO and then outputs permissions for each mailbox into a CSV

    .NOTES
        Version: 0.8
        Updated: 01-07-2022 v0.8    Refactored using a function for efficiency
                                    Fixed issue where non-user trustees (i.e. groups) were not captured
        Updated: 07-01-2022 v0.7    Updated to use .Where method instead of Where-Object for speed
        Updated: 06-01-2022 v0.6    Changed output file date to match order of ISO8601 standard
        Updated: 10-11-2021 v0.5    Disabled write-progress if the verbose parameter is used
        Updated: 08-11-2021 v0.4    Updated filename ordering
        Updated: 18-10-2021 v0.3    Refactored to remove unnecessary lines, add error handling and improve formatting
        Updated: 14-10-2021 v0.2    Updated to use Rest-based commands where possible
        Updated: 01-05-2021 v0.1    Initial draft

        Authors: Luke Allinson (github:LukeAllinson)
                 Robin Dadswell (github:RobinDadswell)

    .PARAMETER OutputPath
        Full path to the folder where the output will be saved.
        Can be used without the parameter name in the first position only.

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

    .PARAMETER IncludeNoPermissions
        Includes mailboxes with no permissions in the export; by default only valid permissions are shown in the export.
        For example, if User01 has Full Access and SendOnBehalf permissions, then only these are shown in the report by default. If the IncludeNoPermissions parameter is included then SendAs permissions will also be included as "<none>".
        Similarly, if User02 has no permissions at all it will not be present in the export, however with this parameter set all three permissions will be included as "<none>".

    .EXAMPLE
        .\Export-EXOMailboxPermissions.ps1 C:\Scripts\
        Exports mailbox permissions for all mailbox types

    .EXAMPLE
        .\Export-EXOMailboxPermissions.ps1 -RecipientTypeDetails RoomMailbox,EquipmentMailbox -OutputPath C:\Scripts\ -IncludeNoPermissions
        Exports mailbox permissions only for Room and Equipment mailboxes; include all permissions, even if blank.

    .EXAMPLE
        .\Export-EXOMailboxPermissions.ps1 C:\Scripts\ -MailboxFilter 'Department -eq "R&D"'
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
            UserPrincipalName           = $mailbox.UserPrincipalName
            DisplayName                 = $mailbox.DisplayName
            PrimarySmtpAddress          = $mailbox.PrimarySmtpAddress
            RecipientTypeDetails        = $mailbox.RecipientTypeDetails
            PermissionType              = $PermissionType
            TrusteeIdentity             = '<NoDefinedPermissions>'
            TrusteeName                 = '<NoDefinedPermissions>'
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
                $permTrustee = $recipients.Where({ ($_.Name -eq $perm.User) -or ($_.PrimarySmtpAddress -eq $perm.User) -or ($_.emailaddresses -contains "smtp:$($faPerm.User)") })
                $trusteeId = $perm.User
            }
            SendAs
            {
                $permTrustee = $recipients.Where({ ($_.Name -eq $perm.Trustee) -or ($_.PrimarySmtpAddress -eq $perm.Trustee) -or ($_.emailaddresses -contains "smtp:$($faPerm.Trustee)") })
                $trusteeId = $perm.Trustee
            }
            SendOnBehalf
            {
                $permTrustee = $recipients.Where({ $_.Name -eq $perm })
                $trusteeId = $perm
            }
        }

        if ($permTrustee)
        {
            $objPermEntry = [ordered]@{
                UserPrincipalName           = $mailbox.UserPrincipalName
                DisplayName                 = $mailbox.DisplayName
                PrimarySmtpAddress          = $mailbox.PrimarySmtpAddress
                RecipientTypeDetails        = $mailbox.RecipientTypeDetails
                PermissionType              = $PermissionType
                TrusteeIdentity             = $permTrustee.PrimarySmtpAddress
                TrusteeName                 = $permTrustee.Name
                TrusteeRecipientTypeDetails = $permTrustee.RecipientTypeDetails
            }
            $output.Add([PSCustomObject]$objPermEntry) | Out-Null
        }
        else
        {
            $objPermEntry = [ordered]@{
                UserPrincipalName           = $mailbox.UserPrincipalName
                DisplayName                 = $mailbox.DisplayName
                PrimarySmtpAddress          = $mailbox.PrimarySmtpAddress
                RecipientTypeDetails        = $mailbox.RecipientTypeDetails
                PermissionType              = $PermissionType
                TrusteeIdentity             = $trusteeId
                TrusteeName                 = '<TrusteeNotFound>'
                TrusteeRecipientTypeDetails = '<TrusteeNotFound>'
            }
            $output.Add([PSCustomObject]$objPermEntry) | Out-Null
        }
    }
    return $output
}

### Main Script
# Check if there is an active Exchange Online PowerShell session and connect if not
$PSSessions = Get-PSSession | Select-Object -Property State, Name
if ((@($PSSessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -eq 0)
{
    Write-Verbose 'Not connected to Exchange Online, prompting to connect'
    Connect-ExchangeOnline
}

# Set Constants and Variables
if (!$PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent)
{
    Write-Progress -Id 1 -Activity 'EXO Mailbox Permissions Report' -Status 'Initialising...' -PercentComplete 15
}
$i = 1
$timeStamp = Get-Date -Format yyyyMMdd-HHmm
$tenantName = (Get-OrganizationConfig).Name.Split('.')[0]
$outputFile = $OutputPath.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar + 'EXOMailboxPermissions_' + $tenantName + '_' + $timeStamp + '.csv'
$output = New-Object System.Collections.Generic.List[System.Object]

# Get In-Scope Mailboxes
$commandHashTable = @{
    Properties  = 'GrantSendOnBehalfTo'
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
    Write-Progress -Id 1 -Activity 'EXO Mailbox Permissions Report' -Status 'Getting In-Scope Mailboxes' -PercentComplete 30
}

Write-Verbose 'Getting Mailboxes from Exchange Online'

try
{
    $mailboxes = @(Get-EXOMailbox @commandHashTable | Sort-Object UserPrincipalName)
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
    Write-Progress -Id 1 -Activity 'EXO Mailbox Permissions Report' -Status 'Getting All Recipients' -PercentComplete 45
}

try
{
    Write-Verbose 'Getting Distribution Groups from Exchange Online'
    $recipients = @(Get-EXORecipient @commandHashTable2 | Sort-Object Name)
}
catch
{
    throw
}

# Get all SendAs permissions
if (!$PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent)
{
    Write-Progress -Id 1 -Activity 'EXO Mailbox Permissions Report' -Status 'Getting All SendAs Permissions' -PercentComplete 60
}

try
{
    Write-Verbose 'Getting Mailbox SendAs permissions from Exchange Online'
    $allSendAsPerms = @(Get-EXORecipientPermission -ResultSize unlimited).Where({ $_.AccessRights -eq 'SendAs' -and $_.Trustee -notmatch 'SELF' })
}
catch
{
    throw
}

#  Loop through the list of mailboxes and output the results to the CSV
if (!$PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent)
{
    Write-Progress -Id 1 -Activity 'EXO Mailbox Permissions Report' -Status 'Processing Mailbox Permissions' -PercentComplete 75
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
        $fullAccessPerms = @(Get-EXOMailboxPermission $mailbox.Identity -ErrorAction stop).Where({ ($_.AccessRights -like 'Full*') -and ($_.User -notmatch 'SELF') })
    }
    catch
    {
        Write-Verbose "Failure getting FullAccess permissions for $($mailbox.UserPrincipalName)"
        $faPermEntry = [ordered]@{
            UserPrincipalName    = $mailbox.UserPrincipalName
            DisplayName          = $mailbox.DisplayName
            PrimarySmtpAddress   = $mailbox.PrimarySmtpAddress
            RecipientTypeDetails = $mailbox.RecipientTypeDetails
            PermissionType       = 'FullAccess'
            TrusteeUPN           = '<ErrorRunningCommand>'
            TrusteeDisplayName   = '<ErrorRunningCommand>'
            TrusteeStatus        = '<ErrorRunningCommand>'
        }
        $output.Add([PSCustomObject]$faPermEntry) | Out-Null
        Continue
    }

    $resolvedFullAccessPerms = [Object[]](Resolve-Permissions -Recipients $recipients -Mailbox $mailbox -Permissions $fullAccessPerms -PermissionType 'FullAccess' -IncludeNoPermissions $IncludeNoPermissions)
    if ($resolvedFullAccessPerms)
    {
        $output.AddRange($resolvedFullAccessPerms)
    }

    # SendAs Permissions
    Write-Verbose "Processing SendAs permissions for $($mailbox.UserPrincipalName)"
    $sendAsPerms = $allSendAsPerms.Where( { $_.Identity -eq $mailbox.Identity } )
    $resolvedSendAsPerms = [Object[]](Resolve-Permissions -Recipients $recipients -Mailbox $mailbox -Permissions $sendAsPerms -PermissionType 'SendAs' -IncludeNoPermissions $IncludeNoPermissions)
    if ($resolvedSendAsPerms)
    {
        $output.AddRange($resolvedSendAsPerms)
    }

    # SendOnBehalf Permissions
    Write-Verbose "Processing SendOnBehalfTo permissions for $($mailbox.UserPrincipalName)"
    $sendOnBehalfPerms = $mailbox.GrantSendOnBehalfTo
    $resolvedSendOnBehalfPerms = [Object[]](Resolve-Permissions -Recipients $recipients -Mailbox $mailbox -Permissions $sendOnBehalfPerms -PermissionType 'SendOnBehalf' -IncludeNoPermissions $IncludeNoPermissions)
    if ($resolvedSendOnBehalfPerms)
    {
        $output.AddRange($resolvedSendOnBehalfPerms)
    }
}

if (!$PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent)
{
    Write-Progress -Id 2 -ParentId 1 -Activity 'Processing Mailbox Permissions' -Completed
    Write-Progress -Id 1 -Activity 'EXO Mailbox Permissions Report' -Status 'Writing Output' -PercentComplete 95
}

$output | Export-Csv $outputFile -NoClobber -NoTypeInformation -Encoding UTF8

# Export to csv
if (!$PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent)
{
    Write-Progress -Id 1 -Activity 'EXO Mailbox Permissions Report' -Completed
}

return "Mailbox permissions have been exported to $outputfile"
