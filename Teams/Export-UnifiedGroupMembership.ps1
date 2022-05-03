#Requires -Version 5 -Modules @{ ModuleName = "ExchangeOnlineManagement"; ModuleVersion = "2.0.5" }


<#
    .SYNOPSIS
        Name: Export-UnifiedGroupMembership.ps1
        This gathers extended Microsoft 365 Group information and exports to a csv file.

    .DESCRIPTION
        This script connects to EXO and then outputs Microsoft 365 Group information and membership to a CSV file.

    .NOTES
        Version: 0.2
        Updated: 19-01-2022  v0.1    Initial draft
        Updated: 03-05-2022  v0.2    Updated ErrorAction to trigger try/catch properly

        Authors: Luke Allinson (github:LukeAllinson)
                 Robin Dadswell (github:RobinDadswell)

    .PARAMETER OutputPath
        Full path to the folder where the output will be saved.
        Can be used without the parameter name in the first position only.

    .PARAMETER IncludeSoftDeletedGroups
        Include soft-deleted Microsoft 365 Groups in the results.

    .PARAMETER IncludeAdditionalAttributes
        Include additional Unified Group information attributes; only basic properties are included by default.

    .PARAMETER GroupFilter
        Provide a filter to reduce the size of the Get-UnifiedGroup query; this must follow oPath syntax standards.
        For example:
        'EmailAddresses -like "*tech*"'
        'DisplayName -like "*wayne*"'

    .PARAMETER Filter
        Alias of GroupFilter parameter.

    .EXAMPLE
        .\Export-UnifiedGroupMembership.ps1 C:\Scripts\ -IncludeAdditionalAttributes
        Exports Microsoft 365 Groups including additional attributes

    .EXAMPLE
        .\Export-UnifiedGroupMembership.ps1 -IncludeSoftDeletedGroups -Output C:\Scripts\
        Exports all Microsoft 365 Group information, including soft-deleted objects

    .EXAMPLE
        .\Export-UnifiedGroupMembership.ps1 C:\Scripts\ -GroupFilter 'Name -like "*R&D*"'
        Exports all Microsoft 365 Groups with "R&D" in the name
#>

[CmdletBinding(DefaultParameterSetName = 'DefaultParameters')]
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
    $IncludeAdditionalAttributes,
    [Parameter()]
    [switch]
    $IncludeSoftDeletedGroups,
    [Parameter()]
    [Alias('Filter')]
    [string]
    $GroupFilter
)

### Main Script
# Check if there is an active Exchange Online PowerShell session and connect if not
$PSSessions = Get-PSSession | Select-Object -Property State, Name
if ((@($PSSessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -eq 0)
{
    Write-Verbose 'Not connected to Exchange Online, prompting to connect'
    Connect-ExchangeOnline
}

# Define constants for use later
$i = 1
$timeStamp = Get-Date -Format yyyyMMdd-HHmm
Write-Verbose 'Getting Tenant Name for file name from Exchange Online'
$tenantName = (Get-OrganizationConfig).Name.Split('.')[0]
$outputFile = $OutputPath.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar + 'EXOUnifiedGroupMembers_' + $tenantName + '_' + $timeStamp + '.csv'
$output = New-Object System.Collections.Generic.List[System.Object]

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

if ($IncludeSoftDeletedGroups)
{
    $commandHashTable['IncludeSoftDeletedGroups'] = $true
}

if ($GroupFilter)
{
    $commandHashTable['Filter'] = $GroupFilter
}

if (!$IncludeAdditionalAttributes)
{
    [System.Collections.ArrayList]$groupProperties = 'Name',
    'DisplayName',
    'PrimarySmtpAddress',
    'Guid',
    'SharePointSiteUrl'
}
else
{
    [System.Collections.ArrayList]$groupProperties = 'Name',
    'DisplayName',
    'PrimarySmtpAddress',
    'Guid',
    'SharePointSiteUrl',
    'AccessType',
    'IsExternalResourcesPublished',
    'AllowAddGuests',
    'WhenSoftDeleted',
    'HiddenFromExchangeClientsEnabled',
    'ExpirationTime',
    'ModerationEnabled',
    'ModeratedBy',
    'GrantSendOnBehalfTo',
    'HiddenFromAddressListsEnabled',
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

# Get M365 Groups using the parameters defined from the hashtable and throw an error if encountered
try
{
    Write-Verbose 'Getting Unified Groups from Exchange Online'
    $unifiedGroups = @(Get-UnifiedGroup @commandHashTable | Select-Object $groupProperties)
}
catch
{
    throw
}

$unifiedGroupsCount = $unifiedGroups.Count
Write-Verbose "There are $unifiedGroupsCount unified groups"

if ($unifiedGroupsCount -eq 0)
{
    return 'There are no unified groups found using the supplied filters'
}

#  Loop through the list of unified groups and output the results to the CSV
Write-Verbose 'Beginning loop through all unified groups'
foreach ($unifiedGroup in $unifiedGroups)
{
    Write-Progress -Id 1 -Activity 'Unified Groups Info' -Status "Processing $($i) of $($unifiedGroupsCount) Unified Groups --- $($unifiedGroup.PrimarySmtpAddress)" -PercentComplete (($i * 100) / $unifiedGroupsCount)

    try
    {
        $unifiedGroupOwners = (Get-UnifiedGroupLinks -Identity $([string]$unifiedGroup.Guid) -LinkType Owners -ResultSize Unlimited -ErrorAction Stop).PrimarySmtpAddress
    }
    # mainly catching authentication timeouts
    catch
    {
        Connect-ExchangeOnline
        $unifiedGroupOwners = (Get-UnifiedGroupLinks -Identity $([string]$unifiedGroup.Guid) -LinkType Owners -ResultSize Unlimited).PrimarySmtpAddress
    }

    try
    {
        $unifiedGroupMembers = (Get-UnifiedGroupLinks -Identity $([string]$unifiedGroup.Guid) -LinkType Members -ResultSize Unlimited -ErrorAction Stop).PrimarySmtpAddress
    }
    # mainly catching authentication timeouts
    catch
    {
        Connect-ExchangeOnline
        $unifiedGroupMembers = (Get-UnifiedGroupLinks -Identity $([string]$unifiedGroup.Guid) -LinkType Members -ResultSize Unlimited).PrimarySmtpAddress
    }

    $unifiedGroupInfo = [ordered]@{
        'Name'               = $unifiedGroup.Name
        'DisplayName'        = $unifiedGroup.DisplayName
        'PrimarySmtpAddress' = $unifiedGroup.PrimarySmtpAddress
        'Guid'               = $unifiedGroup.Guid
        'SharePointSiteUrl'  = $unifiedGroup.SharePointSiteUrl
    }

    if ($IncludeAdditionalAttributes)
    {
        $unifiedGroupInfo['AccessType'] = $unifiedGroup.AccessType
        $unifiedGroupInfo['IsExternalResourcesPublished'] = $unifiedGroup.IsExternalResourcesPublished
        $unifiedGroupInfo['AllowAddGuests'] = $unifiedGroup.AllowAddGuests
        $unifiedGroupInfo['WhenSoftDeleted'] = $unifiedGroup.WhenSoftDeleted
        $unifiedGroupInfo['HiddenFromExchangeClientsEnabled'] = $unifiedGroup.HiddenFromExchangeClientsEnabled
        $unifiedGroupInfo['ExpirationTime'] = $unifiedGroup.ExpirationTime
        $unifiedGroupInfo['ModerationEnabled'] = $unifiedGroup.ModerationEnabled
        $unifiedGroupInfo['ModeratedBy'] = $unifiedGroup.ModeratedBy
        $unifiedGroupInfo['GrantSendOnBehalfTo'] = $unifiedGroup.GrantSendOnBehalfTo
        $unifiedGroupInfo['HiddenFromAddressListsEnabled'] = $unifiedGroup.HiddenFromAddressListsEnabled
        $unifiedGroupInfo['CustomAttribute1'] = $unifiedGroup.CustomAttribute1
        $unifiedGroupInfo['CustomAttribute2'] = $unifiedGroup.CustomAttribute2
        $unifiedGroupInfo['CustomAttribute3'] = $unifiedGroup.CustomAttribute3
        $unifiedGroupInfo['CustomAttribute4'] = $unifiedGroup.CustomAttribute4
        $unifiedGroupInfo['CustomAttribute5'] = $unifiedGroup.CustomAttribute5
        $unifiedGroupInfo['CustomAttribute6'] = $unifiedGroup.CustomAttribute6
        $unifiedGroupInfo['CustomAttribute7'] = $unifiedGroup.CustomAttribute7
        $unifiedGroupInfo['CustomAttribute8'] = $unifiedGroup.CustomAttribute8
        $unifiedGroupInfo['CustomAttribute9'] = $unifiedGroup.CustomAttribute9
        $unifiedGroupInfo['CustomAttribute10'] = $unifiedGroup.CustomAttribute10
        $unifiedGroupInfo['CustomAttribute11'] = $unifiedGroup.CustomAttribute11
        $unifiedGroupInfo['CustomAttribute12'] = $unifiedGroup.CustomAttribute12
        $unifiedGroupInfo['CustomAttribute13'] = $unifiedGroup.CustomAttribute13
        $unifiedGroupInfo['CustomAttribute14'] = $unifiedGroup.CustomAttribute14
        $unifiedGroupInfo['CustomAttribute15'] = $unifiedGroup.CustomAttribute15
        $unifiedGroupInfo['ExtensionCustomAttribute1'] = $unifiedGroup.ExtensionCustomAttribute1
        $unifiedGroupInfo['ExtensionCustomAttribute2'] = $unifiedGroup.ExtensionCustomAttribute2
        $unifiedGroupInfo['ExtensionCustomAttribute3'] = $unifiedGroup.ExtensionCustomAttribute3
        $unifiedGroupInfo['ExtensionCustomAttribute4'] = $unifiedGroup.ExtensionCustomAttribute4
        $unifiedGroupInfo['ExtensionCustomAttribute5'] = $unifiedGroup.ExtensionCustomAttribute5
    }

    $unifiedGroupInfo['OwnersCount'] = $unifiedGroupOwners.Count
    $unifiedGroupInfo['Owners'] = $unifiedGroupOwners -join ';'
    $unifiedGroupInfo['MembersCount'] = $unifiedGroupMembers.Count
    $unifiedGroupInfo['Members'] = $unifiedGroupMembers -join ';'

    $output.Add([PSCustomObject]$unifiedGroupInfo) | Out-Null
    $i++
}

$output | Export-Csv $outputFile -NoClobber -NoTypeInformation -Encoding UTF8
