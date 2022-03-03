#Requires -Version 5 -Modules ExchangeOnlineManagement

<#
    .SYNOPSIS
        Name: Get-EXOMailboxSizes.ps1
        This gathers mailbox size information including primary and archive size and item count.

    .DESCRIPTION
        This script connects to EXO and then gets Mailbox statistics

    .NOTES
        Version: 0.11
        Updated: 01-03-2022 v0.11   Updated to a Get Command - there will be a corresponding Export that utilises this data
        Updated: 01-03-2022 v0.10   Included a paramter to use an input CSV file
        Updated: 06-01-2022 v0.9    Changed output file date to match order of ISO8601 standard
        Updated: 10-11-2021 v0.8    Added parameter sets to prevent use of mutually exclusive parameters
                                    Disabled write-progress if the verbose parameter is used
        Updated: 10-11-2021 v0.7    Updated to include inactive mailboxes and improved error handling
        Updated: 08-11-2021 v0.6    Fixed an issue where archive stats are not included in output if the first mailbox does not have an archive
                                    Updated filename ordering
        Updated: 19-10-2021 v0.5    Updated to use Generic List instead of ArrayList
        Updated: 18-10-2021 v0.4    Updated formatting
        Updated: 15-10-2021 v0.3    Refactored for new parameters, error handling and verbose messaging
        Updated: 14-10-2021 v0.2    Rewritten to improve speed, remove superflous information
        Updated: <unknown>  v0.1    Initial draft

        Authors: Luke Allinson (github:LukeAllinson)
                 Robin Dadswell (github:RobinDadswell)

    .PARAMETER InactiveMailboxOnly
        Only gathers information about inactive mailboxes (active mailboxes are not included in results).

    .PARAMETER IncludeInactiveMailboxes
        Include inactive mailboxes in results; these are not included by default.

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

    .PARAMETER InputCSV
        Full path and filename to an input CSV to specify which mailboxes will be included in the report.
        The CSV must contain a 'UserPrincipalName' or 'PrimarySmtpAddress' or 'EmailAddress' column/header.
        If multiple are found, 'UserPrincipalName' is preferred if found, otherwise 'PrimarySmtpAddress'; 'EmailAddress' is included to cater for exports from non-Exchange (e.g. HR) systems or manually created files.
        Note: All mailboxes are still retrieved and then compared to the CSV to ensure all requested information is captured.
        Note2: Progress is shown as overall progress of all mailboxes plus progress of CSV contents.

    .EXAMPLE
        .\Get-EXOMailboxSizes.ps1
        Gets the size information for all mailbox types

    .EXAMPLE
        .\Get-EXOMailboxSizes.ps1 -RecipientTypeDetails RoomMailbox,EquipmentMailbox
        Gets the size information only for Room and Equipment mailboxes

    .EXAMPLE
        .\Get-EXOMailboxSizes.ps1 -MailboxFilter 'Department -eq "R&D"'
        Gets the size information for all mailboxes from the R&D department
#>

[CmdletBinding(DefaultParameterSetName = 'DefaultParameters')]
param
(
    [Parameter(
        ParameterSetName = 'InactiveOnly'
    )]
    [switch]
    $InactiveMailboxOnly,
    [Parameter(
        ParameterSetName = 'IncludeInactive'
    )]
    [switch]
    $IncludeInactiveMailboxes,
    [Parameter(
        ParameterSetName = 'DefaultParameters'
    )]
    [Parameter(
        ParameterSetName = 'InactiveOnly'
    )]
    [Parameter(
        ParameterSetName = 'IncludeInactive'
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
    [Parameter(
        ParameterSetName = 'InactiveOnly'
    )]
    [Parameter(
        ParameterSetName = 'IncludeInactive'
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
    # Get Mailbox Statistics
    Write-Verbose "Getting mailbox statistics for $($mailbox.PrimarySmtpAddress)"
    try
    {
        $primaryStats = Get-EXOMailboxStatistics -Identity $mailbox.Guid -IncludeSoftDeletedRecipients -Properties LastLogonTime -WarningAction SilentlyContinue -ErrorAction Stop
        $primaryTotalItemSizeMB = $primaryStats | Select-Object @{name = 'TotalItemSizeMB'; expression = { [math]::Round(($_.TotalItemSize.ToString().Split('(')[1].Split(' ')[0].Replace(',', '') / 1MB), 2) } }
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
            $archiveStats = Get-EXOMailboxStatistics -Identity $mailbox.Guid -IncludeSoftDeletedRecipients -Properties LastLogonTime -Archive -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
            $archiveTotalItemSizeMB = $archiveStats | Select-Object @{name = 'TotalItemSizeMB'; expression = { [math]::Round(($_.TotalItemSize.ToString().Split('(')[1].Split(' ')[0].Replace(',', '') / 1MB), 2) } }
        }
        catch
        {
            Write-Error -Message "Error getting archive mailbox statistics for $($mailbox.PrimarySmtpAddress)" -ErrorAction Continue

        }
    }

    # Store everything in an Arraylist
    $mailboxInfo = [ordered]@{
        UserPrincipalName     = $mailbox.UserPrincipalName
        DisplayName           = $mailbox.Displayname
        PrimarySmtpAddress    = $mailbox.PrimarySmtpAddress
        Alias                 = $mailbox.Alias
        RecipientTypeDetails  = $mailbox.RecipientTypeDetails
        IsInactiveMailbox     = $mailbox.IsInactiveMailbox
        LitigationHoldEnabled = $mailbox.LitigationHoldEnabled
        RetentionHoldEnabled  = $mailbox.RetentionHoldEnabled
        InPlaceHolds          = $mailbox.InPlaceHolds -join ';'
        ArchiveStatus         = $mailbox.ArchiveStatus
    }

    if ($primaryStats)
    {
        $mailboxInfo['TotalItemSize(MB)'] = $primaryTotalItemSizeMB.TotalItemSizeMB
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
        $mailboxInfo['Archive_TotalItemSize(MB)'] = $archiveTotalItemSizeMB.TotalItemSizeMB
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
# Check if there is an active Exchange Online PowerShell session and connect if not
$PSSessions = Get-PSSession | Select-Object -Property State, Name
if ((@($PSSessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -eq 0)
{
    Write-Verbose 'Not connected to Exchange Online, prompting to connect'
    Connect-ExchangeOnline
}

# Define constants for use later
$i = 1
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

# Define a hashtable for splatting into Get-EXOMailbox
$commandHashTable = @{
    Properties  = 'LitigationHoldEnabled', 'RetentionHoldEnabled', 'InPlaceHolds', 'ArchiveStatus', 'IsInactiveMailbox'
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
    Write-Verbose 'Getting Mailboxes from Exchange Online'
    $mailboxes = @(Get-EXOMailbox @commandHashTable | Where-Object { $_.RecipientTypeDetails -ne 'DiscoveryMailbox' })
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
        Write-Progress -Id 1 -Activity 'EXO Mailbox Size Report' -Status "Processing $($i) of $($mailboxCount) Mailboxes --- $($mailbox.UserPrincipalName)" -PercentComplete (($i * 100) / $mailboxCount)
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
                Write-Progress -Id 2 -ParentId 1 -Activity 'Mailboxes from CSV' -Status "Processing $($j) of $($csvCount) Mailboxes --- $($mailbox.UserPrincipalName)" -PercentComplete (($j * 100) / $csvCount)
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
        Write-Progress -Activity 'Mailboxes from CSV' -Id 2 -Completed

    }
    Write-Progress -Activity 'EXO Mailbox Size Report' -Id 1 -Completed
}
if ($output)
{
    return $output
}
else
{
    return 'No results returned; no data exported'
}
