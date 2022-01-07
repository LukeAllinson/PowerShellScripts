#Requires -Version 5 -Modules ExchangeOnlineManagement

<#
    .SYNOPSIS
        Name: Export-EXOMailboxPermissions.ps1
        This script enumerates all permissions for every mailbox and exports to a csv file.

    .DESCRIPTION
        This script connects to EXO and then outputs permissions for each mailbox into a CSV

    .NOTES
        Version: 0.7
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

### Main Script
# Check if there is an active Exchange Online PowerShell session and connect if not
$PSSessions = Get-PSSession | Select-Object -Property State, Name
if ((@($PSSessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -eq 0)
{
    Write-Verbose 'Not connected to Exchange Online, prompting to connect'
    Connect-ExchangeOnline
}

# Set Constants and Variables
$i = 1
$timeStamp = Get-Date -Format yyyyMMdd-HHmm
$tenantName = (Get-OrganizationConfig).Name.Split('.')[0]
$outputFile = $OutputPath.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar + 'EXOMailboxPermissions_' + $tenantName + '_' + $timeStamp + '.csv'
$output = New-Object System.Collections.Generic.List[System.Object]

# Define a hashtable for splatting into Get-EXOMailbox
$commandHashTable = @{
    Properties  = 'IsDirSynced', 'GrantSendOnBehalfTo'
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

# Get mailboxes using the parameters defined from the hashtable and throw an error if encountered
try
{
    Write-Verbose 'Getting Mailboxes from Exchange Online'
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

# Get all mailbox SendAs permissions and throw an error if required
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
$i = 1
foreach ($mailbox in $mailboxes)
{
    if (!$PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent)
    {
        Write-Progress -Id 1 -Activity 'EXO Mailbox Permissions Report' -Status "Processing $($i) of $($mailboxCount) Mailboxes --- $($mailbox.UserPrincipalName)" -PercentComplete (($i * 100) / $mailboxCount)
    }

    if (!($mailbox.IsInactiveMailbox))
    {
        try
        {
            # Get Full Access Permissions
            Write-Verbose "Processing FullAccess permissions for $($mailbox.UserPrincipalName)"
            $fullAccessPerms = @(Get-EXOMailboxPermission $mailbox.Identity -ErrorAction stop).Where({ ($_.AccessRights -like 'Full*') -and ($_.User -notmatch 'SELF') })
            if ($fullAccessPerms)
            {
                foreach ($faPerm in $fullAccessPerms)
                {
                    $faUser = $mailboxes.Where( { $_.UserPrincipalName -eq $faPerm.User } )
                    if ($faUser)
                    {
                        $faPermEntry = [ordered]@{
                            UserPrincipalName    = $mailbox.UserPrincipalName
                            DisplayName          = $mailbox.DisplayName
                            PrimarySmtpAddress   = $mailbox.PrimarySmtpAddress
                            RecipientTypeDetails = $mailbox.RecipientTypeDetails
                            IsDirSynced          = $mailbox.IsDirSynced
                            PermissionType       = 'FullAccess'
                            TrusteeUPN           = $faUser.UserPrincipalName
                            TrusteeDisplayName   = $faUser.DisplayName
                        }
                        if ($faUser.IsInactiveMailbox -eq $true)
                        {
                            $faPermEntry['TrusteeStatus'] = 'Inactive'
                        }
                        else
                        {
                            $faPermEntry['TrusteeStatus'] = 'Active'
                        }
                        $output.Add([PSCustomObject]$faPermEntry) | Out-Null
                    }
                }
            }
            elseif ($IncludeNoPermissions)
            {
                $noFAPermEntry = [ordered]@{
                    UserPrincipalName    = $mailbox.UserPrincipalName
                    DisplayName          = $mailbox.DisplayName
                    PrimarySmtpAddress   = $mailbox.PrimarySmtpAddress
                    RecipientTypeDetails = $mailbox.RecipientTypeDetails
                    IsDirSynced          = $mailbox.IsDirSynced
                    PermissionType       = 'FullAccess'
                    TrusteeUPN           = '<none>'
                    TrusteeDisplayName   = $NULL
                    TrusteeStatus        = $NULL
                }
                $output.Add([PSCustomObject]$noFAPermEntry) | Out-Null
            }
        }
        catch
        {
            $faPermEntry = [ordered]@{
                UserPrincipalName    = $mailbox.UserPrincipalName
                DisplayName          = $mailbox.DisplayName
                PrimarySmtpAddress   = $mailbox.PrimarySmtpAddress
                RecipientTypeDetails = $mailbox.RecipientTypeDetails
                IsDirSynced          = $mailbox.IsDirSynced
                PermissionType       = 'FullAccess'
                TrusteeUPN           = 'ERROR RUNNING COMMAND'
                TrusteeDisplayName   = 'ERROR RUNNING COMMAND'
                TrusteeStatus        = 'ERROR RUNNING COMMAND'
            }
            $output.Add([PSCustomObject]$faPermEntry) | Out-Null
        }

        # Get SendAs Permissions
        Write-Verbose "Processing SendAs permissions for $($mailbox.UserPrincipalName)"
        $sendAsPerms = $allSendAsPerms.Where( { $_.Identity -eq $mailbox.Identity } )
        if ($sendAsPerms)
        {
            foreach ($saPerm in $sendAsPerms)
            {
                $saUser = $mailboxes.Where({ $_.UserPrincipalName -eq $saPerm.Trustee })
                if ($saUser)
                {
                    $saPermEntry = [ordered]@{
                        UserPrincipalName    = $mailbox.UserPrincipalName
                        DisplayName          = $mailbox.DisplayName
                        PrimarySmtpAddress   = $mailbox.PrimarySmtpAddress
                        RecipientTypeDetails = $mailbox.RecipientTypeDetails
                        IsDirSynced          = $mailbox.IsDirSynced
                        PermissionType       = 'SendAs'
                        TrusteeUPN           = $saUser.UserPrincipalName
                        TrusteeDisplayName   = $saUser.DisplayName
                    }
                    if ($saUser.IsInactiveMailbox -eq $true)
                    {
                        $saPermEntry['TrusteeStatus'] = 'Inactive'
                    }
                    else
                    {
                        $saPermEntry['TrusteeStatus'] = 'Active'
                    }
                    $output.Add([PSCustomObject]$saPermEntry) | Out-Null
                }
            }
        }
        elseif ($IncludeNoPermissions)
        {
            $noSAPermEntry = [ordered]@{
                UserPrincipalName    = $mailbox.UserPrincipalName
                DisplayName          = $mailbox.DisplayName
                PrimarySmtpAddress   = $mailbox.PrimarySmtpAddress
                RecipientTypeDetails = $mailbox.RecipientTypeDetails
                IsDirSynced          = $mailbox.IsDirSynced
                PermissionType       = 'SendAs'
                TrusteeUPN           = '<none>'
                TrusteeDisplayName   = $NULL
                TrusteeStatus        = $NULL
            }
            $output.Add([PSCustomObject]$noSAPermEntry) | Out-Null
        }

        # Get SendOnBehalf Permissions
        Write-Verbose "Processing SendOnBehalfTo permissions for $($mailbox.UserPrincipalName)"
        $sendOnBehalfPerms = $mailbox.GrantSendOnBehalfTo
        if ($sendOnBehalfPerms)
        {
            foreach ($sobPerms in $sendOnBehalfPerms)
            {
                $sobUser = $mailboxes.Where({ $_.Name -eq $sobPerms })
                if ($sobUser)
                {
                    $sobPermEntry = [ordered]@{
                        UserPrincipalName    = $mailbox.UserPrincipalName
                        DisplayName          = $mailbox.DisplayName
                        PrimarySmtpAddress   = $mailbox.PrimarySmtpAddress
                        RecipientTypeDetails = $mailbox.RecipientTypeDetails
                        IsDirSynced          = $mailbox.IsDirSynced
                        PermissionType       = 'SendOnBehalf'
                        TrusteeUPN           = $sobUsers.UserPrincipalName
                        TrusteeDisplayName   = $sobUsers.DisplayName
                    }
                    if ($sobUser.IsInactiveMailbox -eq $true)
                    {
                        $sobPermEntry['TrusteeStatus'] = 'Inactive'
                    }
                    else
                    {
                        $sobPermEntry['TrusteeStatus'] = 'Active'
                    }
                    $output.Add([PSCustomObject]$sobPermEntry) | Out-Null
                }
            }
        }
        elseif ($IncludeNoPermissions)
        {
            $noSOBPermEntry = [ordered]@{
                UserPrincipalName    = $mailbox.UserPrincipalName
                DisplayName          = $mailbox.DisplayName
                PrimarySmtpAddress   = $mailbox.PrimarySmtpAddress
                RecipientTypeDetails = $mailbox.RecipientTypeDetails
                IsDirSynced          = $mailbox.IsDirSynced
                PermissionType       = 'SendOnBehalf'
                TrusteeUPN           = '<none>'
                TrusteeDisplayName   = $NULL
                TrusteeStatus        = $NULL
            }
            $output.Add([PSCustomObject]$noSOBPermEntry) | Out-Null
        }
        $i++
    }
}

# Export to csv
if (!$PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent)
{
    Write-Progress -Activity 'EXO Mailbox Permissions Report' -Id 1 -Completed
}

$output | Export-Csv $outputFile -NoClobber -NoTypeInformation -Encoding UTF8

return "Mailbox permissions have been exported to $outputfile"
