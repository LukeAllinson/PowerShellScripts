#Requires -Version 5 -Modules ImportExcel

<#
    .SYNOPSIS
        Name: Get-OnPremMailboxSizes.ps1
        This gathers mailbox size information from Exchange On-Premises (2010/2013/2016/2019) including primary and archive size and item count.

    .DESCRIPTION
        This script gets Mailbox statistics

    .NOTES
        Authors: Luke Allinson (github:LukeAllinson)
                 Robin Dadswell (github:RobinDadswell)

    .PARAMETER InactiveMailboxOnly
        Only gathers information about inactive mailboxes (active mailboxes are not included in results).

    .PARAMETER IncludeInactiveMailboxes
        Include inactive mailboxes in results; these are not included by default.

    .PARAMETER RecipientTypeDetails
        Provide one or more RecipientTypeDetails values to return only mailboxes of those types in the results. Seperate multiple values by commas.
        Valid values are: EquipmentMailbox, GroupMailbox, RoomMailbox, SchedulingMailbox, SharedMailbox, TeamMailbox, UserMailbox.

    .PARAMETER MailboxFilter
        Provide a filter to reduce the size of the Get-Mailbox query; this must follow oPath syntax standards.
        For example:
        'EmailAddresses -like "*bruce*"'
        'DisplayName -like "*wayne*"'
        'CustomAttribute1 -eq "InScope"'

    .PARAMETER Filter
        Alias of MailboxFilter parameter.

    .PARAMETER InputCSV
        Full path and filename to an input CSV to specify which mailboxes will be included in the report.
        The CSV must contain a 'UserPrincipalName' or 'PrimarySmtpAddress' or 'EmailAddress' column/header.
        If multiple are found, 'UserPrincipalName' is preferred if found, otherwise 'PrimarySmtpAddress'; 'EmailAddress' is included to cater for exports from non-Exchange (e.g. HR) systems or manually created files.
        Note: All mailboxes are still retrieved and then compared to the CSV to ensure all requested information is captured.
        Note2: Progress is shown as overall progress of all mailboxes plus progress of CSV contents.

    .EXAMPLE
        .\Get-MailboxSizes.ps1
        Gets the size information for all mailbox types

    .EXAMPLE
        .\Get-MailboxSizes.ps1 -RecipientTypeDetails RoomMailbox,EquipmentMailbox
        Gets the size information only for Room and Equipment mailboxes

    .EXAMPLE
        .\Get-MailboxSizes.ps1 -MailboxFilter 'Department -eq "R&D"'
        Gets the size information for all mailboxes from the R&D department
#>

[CmdletBinding(DefaultParameterSetName = 'DefaultParameters')]
param
(
    [Parameter(
        Mandatory,
        Position = 0,
        ParameterSetName = 'DefaultParameters'
    )]
    [Parameter(
        Mandatory,
        Position = 0,
        ParameterSetName = 'InputCSV'
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
    [Parameter(
        ParameterSetName = 'DefaultParameters'
    )]
    [ValidateSet(
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
    [Parameter(
        ParameterSetName = 'DefaultParameters'
    )]
    [Alias('Filter')]
    [string]
    $MailboxFilter,
    [Parameter(
        ParameterSetName = 'InputCSV'
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateScript(
        {
            if (!(Test-Path -Path $_))
            {
                throw "The file $_ does not exist"
            }
            else
            {
                return $true
            }
        }
    )]
    [IO.FileInfo]
    $InputCSV
)
function Get-MailboxInformation ($mailbox)
{
    # Get mailbox Statistics
    Write-Verbose "Getting mailbox statistics for $($mailbox.PrimarySmtpAddress)"
    try
    {
        $primaryStats = Get-MailboxStatistics -Identity $mailbox.Guid.Guid -WarningAction SilentlyContinue -ErrorAction Stop
        $primaryTotalItemSize = $primaryStats | Select-Object @{name = 'TotalItemSize'; expression = { [math]::Round(($_.TotalItemSize.ToString().Split('(')[1].Split(' ')[0].Replace(',', '') / 1MB), 2) } }
    }
    catch
    {
        Write-Error -Message "Error getting mailbox statistics for $($mailbox.PrimarySmtpAddress)" -ErrorAction Continue
    }

    # If an Archive exists, then get Statistics
    if ($mailbox.ArchiveStatus -ne 'None')
    {
        Write-Verbose "Getting archive mailbox statistics for $($mailbox.PrimarySmtpAddress)"
        try
        {
            $archiveStats = Get-MailboxStatistics -Identity $mailbox.Guid.Guid -Archive -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
            $archiveTotalItemSize = $archiveStats | Select-Object @{name = 'TotalItemSize'; expression = { [math]::Round(($_.TotalItemSize.ToString().Split('(')[1].Split(' ')[0].Replace(',', '') / 1MB), 2) } }
        }
        catch
        {
            Write-Error -Message "Error getting archive mailbox statistics for $($mailbox.PrimarySmtpAddress)" -ErrorAction Continue
        }
    }

    # Store everything in an Arraylist
    $mailboxInfo = [ordered]@{
        UserPrincipalName         = $mailbox.UserPrincipalName
        DisplayName               = $mailbox.Displayname
        PrimarySmtpAddress        = $mailbox.PrimarySmtpAddress
        Alias                     = $mailbox.Alias
        SamAccountName            = $mailbox.SamAccountName
        OrganizationalUnit        = $mailbox.OrganizationalUnit
        RecipientTypeDetails      = $mailbox.RecipientTypeDetails
        IsInactiveMailbox         = $mailbox.IsInactiveMailbox
        LitigationHoldEnabled     = $mailbox.LitigationHoldEnabled
        RetentionHoldEnabled      = $mailbox.RetentionHoldEnabled
        InPlaceHolds              = $mailbox.InPlaceHolds -join ';'
        EmailAddressPolicyEnabled = $mailbox.EmailAddressPolicyEnabled
        EmailAddresses            = $mailbox.EmailAddresses -join ';'
        ArchiveStatus             = $mailbox.ArchiveStatus
    }

    if ($primaryStats)
    {
        $mailboxInfo['TotalItemSize(MB)'] = $primaryTotalItemSize.TotalItemSize
        $mailboxInfo['ItemCount'] = $primaryStats.ItemCount
        $mailboxInfo['DeletedItemCount'] = $primaryStats.DeletedItemCount
        $mailboxInfo['LastLogonTime'] = $primaryStats.LastLogonTime
    }
    else
    {
        $mailboxInfo['TotalItemSize(MB)'] = $null
        $mailboxInfo['ItemCount'] = $null
        $mailboxInfo['DeletedItemCount'] = $null
        $mailboxInfo['LastLogonTime'] = $null
    }

    if ($archiveStats)
    {
        $mailboxInfo['Archive_TotalItemSize(MB)'] = $archiveTotalItemSize.TotalItemSize
        $mailboxInfo['Archive_ItemCount'] = $archiveStats.ItemCount
        $mailboxInfo['Archive_DeletedItemCount'] = $archiveStats.DeletedItemCount
        $mailboxInfo['Archive_LastLogonTime'] = $archiveStats.LastLogonTime
    }
    else
    {
        $mailboxInfo['Archive_TotalItemSize(MB)'] = $null
        $mailboxInfo['Archive_ItemCount'] = $null
        $mailboxInfo['Archive_DeletedItemCount'] = $null
        $mailboxInfo['Archive_LastLogonTime'] = $null
    }

    Write-Verbose "Completed gathering mailbox statistics for $($mailbox.PrimarySmtpAddress)"
    return [PSCustomObject]$mailboxInfo
} #End function Get-MailboxInformation

function Compare-EmailAddresses
{
    [OutputType([bool])]
    param
    (
        [Parameter(Mandatory)]
        [System.Collections.Generic.List[System.Object]]
        $EmailAddresses,
        [Parameter(Mandatory)]
        [System.Array]
        $CsvValues
    )
    Write-Verbose 'Comparing column to EmailAddresses'
    foreach ($emailAddress in $EmailAddresses)
    {
        $strippedAddress = $emailAddress.Split(':')[1]
        if ($strippedAddress -in $CsvValues)
        {
            return $true
        }
    }
    return $false

} #End function Compare-EmailAddresses

### Main Script
# Define constants for use later
$i = 1
$timeStamp = Get-Date -Format yyyyMMdd-HHmm
$outputFile = $OutputPath.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar + 'MailboxSizes' + '_' + $timeStamp + '.xlsx'
$output = New-Object System.Collections.Generic.List[System.Object]

# Import and validate inputCSV if specified
if ($InputCSV)
{
    $csv = Import-Csv $InputCSV -Delimiter ','
    $csvHeaders = ($csv | Get-Member -MemberType NoteProperty).Name.ToLower()
    if ('userprincipalname' -notin $csvHeaders -and 'emailaddress' -notin $csvHeaders -and 'primarysmtpaddress' -notin $csvHeaders)
    {
        throw "The file $InputCSV is invalid; cannot find the 'UserPrincipalName', 'Emailaddress' or 'PrimarySmtpAddress' column headings.`
            Please ensure the CSV contains at least one of these headings."
    }
    $csvCount = $csv.Count
    Write-Verbose "There are $csvCount mailboxes in the InputCSV file $InputCSV"
    if ($csvCount -eq 0)
    {
        return 'There are no mailboxes found in the InputCSV file $InputCSV'
    }
    ## create new variable to contain column we are going to use
    # all 3 headers supplied
    if ('userprincipalname' -in $csvHeaders -and 'emailaddress' -in $csvHeaders -and 'primarysmtpaddress' -in $csvHeaders)
    {
        $csvCompare = $csv.userprincipalname
        Write-Verbose '3 columns supplied; using primarysmtpaddress'
    }
    # userprincipalname and emailaddress
    elseif ('userprincipalname' -in $csvHeaders -and 'emailaddress' -in $csvHeaders)
    {
        $csvCompare = $csv.userprincipalname
        Write-Verbose 'userprincipalname and emailaddress columns supplied; using emailaddress'
    }
    # userprincipalname and primarysmtpaddress
    elseif ('userprincipalname' -in $csvHeaders -and 'primarysmtpaddress' -in $csvHeaders)
    {
        $csvCompare = $csv.userprincipalname
        Write-Verbose 'userprincipalname and primarysmtpaddress columns supplied; using primarysmtpaddress'
    }
    # emailaddress and primarysmtpaddress
    elseif ('emailaddress' -in $csvHeaders -and 'primarysmtpaddress' -in $csvHeaders)
    {
        $csvCompare = $csv.primarysmtpaddress
        Write-Verbose 'emailaddress and primarysmtpaddress columns supplied; using primarysmtpaddress'
    }
    # only userprincipalname
    elseif ('userprincipalname' -in $csvHeaders)
    {
        $csvCompare = $csv.userprincipalname
        Write-Verbose 'only userprincipalname column supplied; using userprincipalname'
    }
    # only emailaddress
    elseif ('emailaddress' -in $csvHeaders)
    {
        $csvCompare = $csv.emailaddress
        Write-Verbose 'only emailaddress column supplied; using emailaddress'
    }
    # only primarysmtpaddress
    elseif ('primarysmtpaddress' -in $csvHeaders)
    {
        $csvCompare = $csv.primarysmtpaddress
        Write-Verbose 'only primarysmtpaddress column supplied; using emailaddress'
    }
    $j = 1
}

# Define a hashtable for splatting into Get-Mailbox
$commandHashTable = @{
    ResultSize  = 'Unlimited'
    ErrorAction = 'Stop'
}

if ($IncludeInactiveMailboxes)
{
    $commandHashTable['IncludeInactiveMailbox'] = $true
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
    Write-Verbose 'Getting mailboxes from Exchange'
    $mailboxes = @(Get-Mailbox @commandHashTable | Where-Object { $_.RecipientTypeDetails -ne 'DiscoveryMailbox' })
}
catch
{
    throw
}

$mailboxCount = $mailboxes.Count
Write-Verbose "There are $mailboxCount mailboxes"

if ($mailboxCount -eq 0)
{
    throw 'There are no mailboxes found using the supplied filters'
}

#  Loop through the list of mailboxes and output the results to the CSV
Write-Verbose 'Beginning loop through all mailboxes'
foreach ($mailbox in $mailboxes)
{
    if (!$PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent)
    {
        Write-Progress -Id 1 -Activity 'Getting mailboxes from Exchange' -Status "Processing $($i) of $($mailboxCount) mailboxes --- $($mailbox.UserPrincipalName)" -PercentComplete (($i * 100) / $mailboxCount)
    }

    # if InputCSV is specified, match against mailbox list
    if ($InputCSV)
    {
        if ($j -gt $csvCount)
        {
            Write-Verbose 'All CSV mailboxes found; exiting foreach loop'
            break
        }
        if ($mailbox.UserPrincipalName -in $csvCompare -or $mailbox.PrimarySmtpAddress -in $csvCompare -or (Compare-EmailAddresses -EmailAddresses $mailbox.EmailAddresses -CsvValues $csvCompare))
        {
            if (!$PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent)
            {
                Write-Progress -Id 2 -ParentId 1 -Activity 'Processed mailboxes from csv' -Status "Processing $($j) of $($csvCount)" -PercentComplete (($j * 100) / $csvCount)
            }
            $mailboxInfo = Get-MailboxInformation $mailbox
            $output.Add([PSCustomObject]$mailboxInfo) | Out-Null
            $j++
        }
    }
    else
    {
        $mailboxInfo = Get-MailboxInformation $mailbox
        $output.Add([PSCustomObject]$mailboxInfo) | Out-Null
    }
    $i++
}

if (!$PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent)
{
    if ($InputCSV)
    {
        Write-Progress -Activity 'Processed mailboxes from csv' -Id 2 -Completed

    }
    Write-Progress -Activity 'Getting mailboxes from Exchange' -Id 1 -Completed
}

if ($output.Count -ge 1)
{
    $output | Export-Excel -Path $outputFile -WorksheetName 'MailboxStats' -FreezeTopRow -AutoSize -AutoFilter -TableName 'MailboxStats'

    ### Add summary sheet and apply formatting
    $summaryPage = New-Object System.Collections.Generic.List[System.Object]
    $mailboxTypes = ('UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox')
    $t = 3
    foreach ($mailboxType in $mailboxTypes)
    {
        $summaryStats = [ordered]@{
            'MailboxType'             = $mailboxType
            'MailboxCount'            = "=COUNTIF(MailboxStats[RecipientTypeDetails],A$t)"
            'TotalSize(MB)'           = "=SUMIF(MailboxStats[RecipientTypeDetails],A$t,MailboxStats[TotalItemSize(MB)])"
            'AverageSize(MB)'         = "=IF(C$t<>0,(AVERAGEIF(MailboxStats[RecipientTypeDetails],A$t,MailboxStats[TotalItemSize(MB)])),0)"
            'TotalItemCount'          = "=SUMIF(MailboxStats[RecipientTypeDetails],A$t,MailboxStats[ItemCount])"
            'AverageItemCount'        = "=IF(E$t<>0,(AVERAGEIF(MailboxStats[RecipientTypeDetails],A$t,MailboxStats[TotalItemSize(MB)])),0)"
            'ArchiveCount'            = "=COUNTIFS(MailboxStats[RecipientTypeDetails],A$t,MailboxStats[ArchiveStatus],""Active"")"
            'ArchiveTotalSize(MB)'    = "=SUMIF(MailboxStats[RecipientTypeDetails],A$t,MailboxStats[Archive_TotalItemSize(MB)])"
            'ArchiveAverageSize(MB)'  = "=IF(H$t<>0,(AVERAGEIF(MailboxStats[RecipientTypeDetails],A$t,MailboxStats[Archive_TotalItemSize(MB)])),0)"
            'ArchiveTotalItemCount'   = "=SUMIF(MailboxStats[RecipientTypeDetails],A$t,MailboxStats[Archive_ItemCount])"
            'ArchiveAverageItemCount' = "=IF(J$t<>0,(AVERAGEIF(MailboxStats[RecipientTypeDetails],A$t,MailboxStats[Archive_ItemCount])),0)"
        }
        $summaryPage.Add([PSCustomObject]$summaryStats) | Out-Null
        $t++
    }
    $summaryStats = [ordered]@{
        'MailboxType'             = 'Total'
        'MailboxCount'            = '=SUM(B3:B6)'
        'TotalSize(MB)'           = '=SUM(C3:C6)'
        'AverageSize(MB)'         = '=SUM(D3:D6)'
        'TotalItemCount'          = '=SUM(E3:E6)'
        'AverageItemCount'        = '=SUM(F3:F6)'
        'ArchiveCount'            = '=SUM(G3:G6)'
        'ArchiveTotalSize(MB)'    = '=SUM(H3:H6)'
        'ArchiveAverageSize(MB)'  = '=SUM(I3:I6)'
        'ArchiveTotalItemCount'   = '=SUM(J3:J6)'
        'ArchiveAverageItemCount' = '=SUM(K3:K6)'
    }
    $summaryPage.Add([PSCustomObject]$summaryStats) | Out-Null
    $summaryPage | Export-Excel -Path $outputFile -WorksheetName 'Summary' -TableName 'Summary' -AutoSize -StartRow 2 -MoveToStart
    $output | Sort-Object 'TotalItemSize(MB)' -Descending | Select-Object UserPrincipalName, DisplayName, RecipientTypeDetails, 'TotalItemSize(MB)', ItemCount -First 10 | Export-Excel -Path $outputFile -WorksheetName 'Summary' -TableName 'Top10BySize' -StartRow 10
    $output | Sort-Object ItemCount -Descending | Select-Object UserPrincipalName, DisplayName, RecipientTypeDetails, 'TotalItemSize(MB)', ItemCount -First 10 | Export-Excel -Path $outputFile -WorksheetName 'Summary' -TableName 'Top10ByItemCount' -StartRow 23
    $output | Sort-Object 'Archive_TotalItemSize(MB)' -Descending | Select-Object UserPrincipalName, DisplayName, RecipientTypeDetails, 'Archive_TotalItemSize(MB)', Archive_ItemCount -First 10 | Export-Excel -Path $outputFile -WorksheetName 'Summary' -TableName 'Top10ByArchiveSize' -StartRow 36
    $output | Sort-Object Archive_ItemCount -Descending | Select-Object UserPrincipalName, DisplayName, RecipientTypeDetails, 'Archive_TotalItemSize(MB)', Archive_ItemCount -First 10 | Export-Excel -Path $outputFile -WorksheetName 'Summary' -TableName 'Top10ByArchiveItemCount' -StartRow 49
    $excelPkg = Open-ExcelPackage -Path $outputFile
    $summarySheet = $excelPkg.Workbook.Worksheets['Summary']
    $summarySheet.Cells[1, 1].Value = 'Summary'
    $summarySheet.Cells[9, 1].Value = 'Top10 Mailboxes By Size'
    $summarySheet.Cells[22, 1].Value = 'Top10 Mailboxes By ItemCount'
    $summarySheet.Cells[35, 1].Value = 'Top10 Archives By Size'
    $summarySheet.Cells[48, 1].Value = 'Top10 Archives By ItemCount'
    $summarySheet.Select('A1')
    $summarySheet.SelectedRange.Style.Font.Size = 16
    $summarySheet.SelectedRange.Style.Font.Bold = $true
    $summarySheetTitleCells = ('A9', 'A22', 'A35', 'A48')
    ForEach ($cell in $summarySheetTitleCells)
    {
        $summarySheet.Select($cell)
        $summarySheet.SelectedRange.Style.Font.Size = 13
        $summarySheet.SelectedRange.Style.Font.Bold = $true
    }
    $summarySheetBoldRanges = ('A7:K7', 'D11:D20', 'E24:E33', 'D37:D46', 'E50:E59')
    ForEach ($range in $summarySheetBoldRanges)
    {
        $summarySheet.Select($range)
        $summarySheet.SelectedRange.Style.Font.Bold = $true
    }
    $summarySheetAverageRanges = ('D3:D7', 'F3:F7', 'I3:I7', 'K3:K7')
    ForEach ($range in $summarySheetAverageRanges)
    {
        $summarySheet.Select($range)
        $summarySheet.SelectedRange.Style.Numberformat.Format = '0.00'
    }
    $fullRange = ($summarySheet.Dimension | Select-Object Address).Address
    $summarySheet.Select($fullRange)
    $summarySheet.SelectedRange.AutoFitColumns()
    $summarySheet.Select('A1')
    Close-ExcelPackage $excelPkg

    return "Mailbox size data has been exported to $outputfile"
}
else
{
    return 'No results returned; no data exported'
}
