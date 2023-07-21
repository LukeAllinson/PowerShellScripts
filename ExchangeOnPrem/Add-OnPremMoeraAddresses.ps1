<#
    .SYNOPSIS
        Name: Add-OnPremMoeraAddresses.ps1
        Adds Microsoft Online Exchange Routing Addresses (MOERA) to Exchange On-Premises (2010/2013/2016/2019) mailboxes, required for migration to Exchange Online.

    .DESCRIPTION
        In an Exchange Hybrid Configuration, any Mailboxes moving to Exchange Online must have a Microsoft Online Exchange Routing Address (MOERA).
        This script adds the required addresses to either any mailbox that does not have one, or to mailboxes specified in a CSV.
        The required moera address is determined by searching for an Accepted Domain with the format <something>.mail.onmicrosoft.com. If no Accepted Domains or multiple Accepted Domains of this type are found, the script will exit.

    .NOTES
        Authors: Luke Allinson (github:LukeAllinson)
                 Robin Dadswell (github:RobinDadswell)

    .PARAMETER OutputPath
        Full path to a folder to save the results report.

    .PARAMETER InputCSV
        Full path and filename to an input CSV to specify which mailboxes will be processed; if a CSV is not specified then all mailboxes will be processed.
        The CSV must contain a 'UserPrincipalName' or 'PrimarySmtpAddress' or 'EmailAddress' column/header.
        If multiple are found, 'UserPrincipalName' is preferred if found, otherwise 'PrimarySmtpAddress'; 'EmailAddress' is included to cater for exports from non-Exchange (e.g. HR) systems or manually created files.
        Note: All mailboxes are still retrieved and then compared to the CSV to ensure all requested information is captured.
        Note2: Progress is shown as overall progress of all mailboxes plus progress of CSV contents.

    .PARAMETER ReportOnly
        Runs the Set-Mailbox command with the "WhatIf" parameter, meaning that no actual changes are made but the command is tested to ensure there are no errors and a report is still generated.

    .EXAMPLE
        .\Add-OnPremMoeraAddresses.ps1 -OutputPath C:\Scripts\
        Adds Moera addresses to all mailboxes that do not have one.

    .EXAMPLE
        .\Add-OnPremMoeraAddresses.ps1 -OutputPath C:\Scripts\ -InputCSV C:\Files\Mailboxes.csv
        Adds Moera addresses to mailboxes specified in the CSV file.
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
    $InputCSV,
    [Parameter()]
    [switch]
    $ReportOnly
)
function Add-MoeraAddress
{
    param
    (
        [Parameter(Mandatory)]
        [System.Object]
        $Mailbox,
        [Parameter(Mandatory)]
        [string]
        $MoeraAddressDomain,
        [Parameter(Mandatory)]
        [System.Boolean]
        $ReportOnly
    )
    $moeraAddress = [string]::Join('@', $Mailbox.Alias, $MoeraAddressDomain)
    $moeraOutput = [ordered]@{
        Identity             = $Mailbox.Identity
        UserPrincipalName    = $Mailbox.UserPrincipalName
        DisplayName          = $Mailbox.Displayname
        PrimarySmtpAddress   = $Mailbox.PrimarySmtpAddress
        Alias                = $Mailbox.Alias
        RecipientTypeDetails = $Mailbox.RecipientTypeDetails
        MoeraAddress         = $moeraAddress
    }
    try
    {
        if ($ReportOnly)
        {
            Set-Mailbox -Identity $Mailbox.Identity -EmailAddresses @{ add = $moeraAddress } -ErrorAction Stop -WhatIf
            $moeraOutput['AddedMoeraAddress'] = 'ReportOnly: No changes made'
            $moeraOutput['Error'] = ''
        }
        else
        {
            Set-Mailbox -Identity $Mailbox.Identity -EmailAddresses @{ add = $moeraAddress } -ErrorAction Stop
            $moeraOutput['AddedMoeraAddress'] = 'Success'
            $moeraOutput['Error'] = ''
        }
        Write-Verbose "Successfully added $moeraAddress to mailbox $($Mailbox.UserPrincipalName)"
    }
    catch
    {
        $moeraOutput['AddedMoeraAddress'] = 'Fail'
        $moeraOutput['Error'] = $_.Exception.Message
        Write-Error "Failed to add $moeraAddress to mailbox $($Mailbox.UserPrincipalName)" -Category
    }
    return [PSCustomObject]$moeraOutput
} # end function Add-MoeraAddress

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
        if ($emailAddress -in $CsvValues)
        {
            return $true
        }
    }
    return $false
} #End function Compare-EmailAddresses

# Define constants for use later
$i = 1
$timeStamp = Get-Date -Format yyyyMMdd-HHmm
$outputFile = $OutputPath.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar + 'Add-OnPremMoeraAddresses_Results_' + $timeStamp + '.csv'
$output = New-Object System.Collections.Generic.List[System.Object]

# Import and validate inputCSV if specified
if ($InputCSV)
{
    $csv = Import-Csv $InputCSV -Delimiter ','
    $csvCount = $csv.Count
    Write-Verbose "There are $csvCount mailboxes in the InputCSV file $InputCSV"
    if ($csvCount -eq 0)
    {
        return "There are no mailboxes found in the InputCSV file $InputCSV"
    }
    $csvHeaders = ($csv | Get-Member -MemberType NoteProperty).Name.ToLower()
    if ('userprincipalname' -notin $csvHeaders -and 'emailaddress' -notin $csvHeaders -and 'primarysmtpaddress' -notin $csvHeaders)
    {
        throw "The file $InputCSV is invalid; cannot find the 'UserPrincipalName', 'Emailaddress' or 'PrimarySmtpAddress' column headings.`
            Please ensure the CSV contains at least one of these headings."
    }
    ## create new variable to contain column we are going to use
    # all 3 headers supplied
    if ('userprincipalname' -in $csvHeaders -and 'emailaddress' -in $csvHeaders -and 'primarysmtpaddress' -in $csvHeaders)
    {
        $csvCompare = $csv.userprincipalname
        Write-Verbose '3 columns supplied; using userprincipalname'
    }
    # userprincipalname and emailaddress
    elseif ('userprincipalname' -in $csvHeaders -and 'emailaddress' -in $csvHeaders)
    {
        $csvCompare = $csv.userprincipalname
        Write-Verbose 'userprincipalname and emailaddress columns supplied; using userprincipalname'
    }
    # userprincipalname and primarysmtpaddress
    elseif ('userprincipalname' -in $csvHeaders -and 'primarysmtpaddress' -in $csvHeaders)
    {
        $csvCompare = $csv.userprincipalname
        Write-Verbose 'userprincipalname and primarysmtpaddress columns supplied; using userprincipalname'
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
        Write-Verbose 'only primarysmtpaddress column supplied; using primarysmtpaddress'
    }
    $j = 1
}

# Find the MOERA address and throw script if more than 1 is found
$moeraDomain = Get-AcceptedDomain *.mail.onmicrosoft.com
if ($null -eq $moeraDomain)
{
    throw 'No MOERA (.mail.onmicrosoft.com) domain found. Exiting script.'
}
if ($moeraDomain.count -gt 1)
{
    foreach ($domain in $moeraDomain)
    {
        '{0} - {1}' -f ($moeraDomain.IndexOf($domain) + 1), $domain.DomainName.Address
    }
    $Choice = ''
    while ([string]::IsNullOrEmpty($Choice))
    {
        $Choice = Read-Host 'Please choose domain by number '
        if ($Choice -notin 1..$moeraDomain.Count)
        {
            [console]::Beep(1000, 300)
            Write-Warning ''
            # Robin's favourite joke code snippet
            Write-Warning ('    Your choice [ {0} ] is not valid.' -f $Choice)
            Write-Warning ('        The valid choices are 1 thru {0}.' -f $moeraDomain.Count)
            Write-Warning '        Please try again ...'
            $Choice = ''
        }
    }
    $moeraDomain = $moeraDomain[$Choice - 1]
}
$moeraAddressDomain = $moeraDomain.DomainName.Address

# Define a hashtable for splatting into Get-Mailbox
$commandHashTable = @{
    ResultSize  = 'Unlimited'
    ErrorAction = 'Stop'
    Filter      = "emailAddresses -notlike `"*@$moeraAddressDomain`" -and RecipientTypeDetails -ne 'DiscoveryMailbox'"
}

try
{
    Write-Verbose 'Getting mailboxes from Exchange'
    $mailboxes = @(Get-Mailbox @commandHashTable)
}
catch
{
    throw
}
$mailboxCount = $mailboxes.Count
Write-Verbose "There are $mailboxCount mailboxes"

Write-Verbose 'Beginning loop through all mailboxes'
foreach ($mailbox in $mailboxes)
{
    if (!$PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent)
    {
        Write-Progress -Id 1 -Activity 'Add On-Premises MOERA Addresses' -Status "Processing $($i) of $($mailboxCount) mailboxes --- $($mailbox.UserPrincipalName)" -PercentComplete (($i * 100) / $mailboxCount)
    }

    # if InputCSV is specified, match against mailbox list
    if ($InputCSV)
    {
        Write-Verbose 'Using InputCSV'
        if ($j -gt $csvCount)
        {
            Write-Verbose 'All CSV mailboxes found; exiting foreach loop'
            break
        }
        if ($mailbox.UserPrincipalName -in $csvCompare -or $mailbox.PrimarySmtpAddress -in $csvCompare -or (Compare-EmailAddresses -EmailAddresses $mailbox.EmailAddresses.SmtpAddress -CsvValues $csvCompare))
        {
            if (!$PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent)
            {
                Write-Progress -Id 2 -ParentId 1 -Activity 'Processed mailboxes from csv' -Status "Processing $($j) of $($csvCount)" -PercentComplete (($j * 100) / $csvCount)
            }
            $addMoeraAddress = Add-MoeraAddress -Mailbox $mailbox -MoeraAddressDomain $moeraAddressDomain -ReportOnly $ReportOnly
            $output.Add([PSCustomObject]$addMoeraAddress) | Out-Null
            $j++
        }
    }
    else
    {
        $addMoeraAddress = Add-MoeraAddress -Mailbox $mailbox -MoeraAddressDomain $moeraAddressDomain -ReportOnly $ReportOnly
        $output.Add([PSCustomObject]$addMoeraAddress) | Out-Null
    }
    $i++
}

if (!$PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent)
{
    if ($InputCSV)
    {
        Write-Progress -Activity 'Processed mailboxes from csv' -Id 2 -Completed
    }
    Write-Progress -Activity 'Getting mailboxes from Exchange Online' -Id 1 -Completed
}
if ($output)
{
    $output | Export-Csv $outputFile -NoClobber -NoTypeInformation -Encoding UTF8
}
else
{
    return 'No results returned; no data exported'
}
