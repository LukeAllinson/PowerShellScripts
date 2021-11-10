#Requires -Version 5 -Modules ExchangeOnlineManagement

<#
    .SYNOPSIS
        Name: Export-EXOPublicFolderSizes.ps1
        This gathers Public Folder information and sizes and exports to a csv file.

    .DESCRIPTION
        This script connects to EXO and then outputs Public Folder information to a CSV file.

    .NOTES
        Version: 0.3
        Updated: 08-11-2021 v0.3    Updated filename ordering
        Updated: 19-10-2021 v0.2    Refactored using current script standards
        Updated: <unknown>  v0.1    Initial draft

        Authors: Luke Allinson (github:LukeAllinson)
                 Robin Dadswell (github:RobinDadswell)

    .PARAMETER OutputPath
        Full path to the folder where the output will be saved.
        Can be used without the parameter name in the first position only.

    .EXAMPLE
        .\Export-EXOPublicFolderSizes.ps1 C:\Scripts\
        Exports all Public Folder information to a csv file in C:\Scripts
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
    $OutputPath
)

function Get-PublicFolderInformation ($publicFolder)
{
    # Get Public Folder Statistics
    try
    {
        Write-Verbose "Getting Public Folder statistics for $($publicFolder.identity)"
        $publicFolderStats = Get-PublicFolderStatistics -Identity $publicFolder.identity -WarningAction SilentlyContinue -ErrorAction Stop
    }
    catch
    {
        throw
    }

    if ($publicFolderStats.OwnerCount -ne 0)
    {
        try
        {
            $publicFolderOwners = (Get-PublicFolderClientPermission -Identity $publicFolder.Identity | Where-Object { $_.AccessRights -eq 'Owner' } | Select-Object User).DisplayName -join ';'
        }
        catch
        {
            throw
        }
    }
    $TotalItemSizeMB = $publicFolderStats | Select-Object @{name = 'TotalItemSizeMB'; expression = { [math]::Round(($_.TotalItemSize.ToString().Split('(')[1].Split(' ')[0].Replace(',', '') / 1MB), 2) } }

    # Store everything in an Arraylist
    $publicFolderInfo = [ordered]@{
        Identity             = $publicFolder.Identity
        Name                 = $publicFolder.Name
        FolderClass          = $publicFolder.FolderClass
        MailEnabled          = $publicFolder.MailEnabled
        ContentMailboxName   = $publicFolder.ContentMailboxName
        ItemCount            = $publicFolderStats.ItemCount
        'TotalItemSize(MB)'  = $TotalItemSizeMB.TotalItemSizeMB
        ContactCount         = $publicFolderStats.ContactCount
        OwnerCount           = $publicFolderStats.OwnerCount
        Owners               = $publicFolderOwners
        CreationTime         = $publicFolderStats.CreationTime
        LastModificationTime = $publicFolderStats.LastModificationTime
    }

    Write-Verbose "Completed gathering Public Folder statistics for $($publicFolder.PrimarySmtpAddress)"
    return [PSCustomObject]$publicFolderInfo
} #End Function Get-PublicFolderInformation

### Main Script
# Check if there is an active Exchange Online PowerShell session and connect if not
$PSSessions = Get-PSSession | Select-Object -Property State, Name
if ((@($PSSessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -eq 0)
{
    Write-Verbose 'Not connected to Exchange Online, prompting to connect'
    Connect-ExchangeOnline
}

# Define constants for use later
$timeStamp = Get-Date -Format ddMMyyyy-HHmm
Write-Verbose 'Getting Tenant Name for file name from Exchange Online'
$tenantName = (Get-OrganizationConfig).Name.Split('.')[0]
$outputFile = $OutputPath.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar + 'EXOPublicFolderSizes_' + $tenantName + '_' + $timeStamp + '.csv'
$output = New-Object System.Collections.Generic.List[System.Object]

Write-Verbose "Checking if $outputFile already exists"
if (Test-Path $outputFile -ErrorAction SilentlyContinue)
{
    throw "The file $outputFile already exists, please delete the file and try again."
}

# Define a hashtable for splatting into Get-PublicFolder
$commandHashTable = @{
    Recurse     = $true
    ResultSize  = 'Unlimited'
    ErrorAction = 'Stop'
}

# Get Public Folders using the parameters defined from the hashtable and throw an error if encountered
try
{
    Write-Verbose 'Getting Public Folders from Exchange Online'
    $publicFolders = @(Get-PublicFolder @commandHashTable)
}
catch
{
    throw
}

$publicFolderCount = $publicFolders.Count
Write-Verbose "There are $publicFolderCount Public Folders"

if ($publicFolderCount -eq 0)
{
    return 'There are no Public Folders found'
}

#  Loop through the list of Public Folders and output the results to the CSV
Write-Verbose 'Beginning loop through all Public Folders'
foreach ($publicFolder in $publicFolders)
{
    Write-Progress -Id 1 -Activity 'EXO Public Folder Size Report' -Status "Processing $($i) of $($publicFolderCount) Public Folders --- $($publicFolder.Identity)" -PercentComplete (($i * 100) / $publicFolderCount)
    try
    {
        $publicFolderInfo = Get-PublicFolderInformation $publicFolder
        $output.Add([PSCustomObject]$publicFolderInfo) | Out-Null
    }
    catch
    {
        continue
    }
    finally
    {
        $i++
    }
}

Write-Progress -Activity 'EXO Public Folder Size Report' -Id 1 -Completed
$output | Export-Csv $outputFile -NoClobber -NoTypeInformation -Encoding UTF8

return "Public Folder information and size data has been exported to $outputfile"
