<#
    .SYNOPSIS
        Name: Remove-OnPremLegacyEmailAddresses.ps1
        Removes addresses using a specified email Domain from Exchange On-Premises (2013/2016/2019) mailboxes.

    .DESCRIPTION
        This script removes any instances of email addresses using a specified Domain from Exchange On-Premises (2013/2016/2019) mailboxes.
        This may be useful for tidy-up operations prior to a Hybrid mailbox migration, or if an Accepted Domain is removed.

    .NOTES
        Authors: Luke Allinson (github:LukeAllinson)
                 Robin Dadswell (github:RobinDadswell)

    .PARAMETER DomainName
        Specifies the Domain for which email addresses will be removed from mailboxes.

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
    [ValidatePattern('(?=^.{1,253}$)(^((?!-)[a-zA-Z0-9-]{1,63}(?<!-)\.)+[a-zA-Z]{2,63}$)')]
    [string]
    $DomainName,
    [Parameter(
        Mandatory
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
function Remove-LegacyAddress
{
    param
    (
        [Parameter(Mandatory)]
        [System.Object]
        $Mailbox,
        [Parameter(Mandatory)]
        [string]
        $LegacyAddress,
        [Parameter(Mandatory)]
        [System.Boolean]
        $ReportOnly
    )
    $legacyAddressOutput = [ordered]@{
        Identity             = $Mailbox.Identity
        UserPrincipalName    = $Mailbox.UserPrincipalName
        DisplayName          = $Mailbox.Displayname
        PrimarySmtpAddress   = $Mailbox.PrimarySmtpAddress
        Alias                = $Mailbox.Alias
        RecipientTypeDetails = $Mailbox.RecipientTypeDetails
        LegacyAddress        = $LegacyAddress
    }
    try
    {
        if ($ReportOnly)
        {
            Set-Mailbox -Identity $Mailbox.Identity -EmailAddresses @{ Remove = $LegacyAddress } -ErrorAction Stop -WhatIf
            $legacyAddressOutput['RemovedLegacyAddress'] = 'ReportOnly: No changes made'
            $legacyAddressOutput['Error'] = ''
        }
        else
        {
            Set-Mailbox -Identity $Mailbox.Identity -EmailAddresses @{ Remove = $LegacyAddress } -ErrorAction Stop
            $legacyAddressOutput['RemovedLegacyAddress'] = 'Success'
            $legacyAddressOutput['Error'] = ''
        }
        Write-Verbose "Successfully added $LegacyAddress to mailbox $($Mailbox.UserPrincipalName)"
    }
    catch
    {
        $legacyAddressOutput['RemovedLegacyAddress'] = 'Fail'
        $legacyAddressOutput['Error'] = $_.Exception.Message
        Write-Error "Failed to add $LegacyAddress to mailbox $($Mailbox.UserPrincipalName)" -Category
    }
    return [PSCustomObject]$legacyAddressOutput
} # end function Remove-LegacyAddress

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
$outputFile = $OutputPath.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar + 'Remove-OnPremLegacyEmailAddresses_Results_' + $timeStamp + '.csv'
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

# Define a hashtable for splatting into Get-Mailbox
$commandHashTable = @{
    ResultSize  = 'Unlimited'
    ErrorAction = 'Stop'
    Filter      = "emailAddresses -like `"*@$DomainName`" -and RecipientTypeDetails -ne 'DiscoveryMailbox'"
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
        Write-Progress -Id 1 -Activity 'Remove Legacy Email Addresses' -Status "Processing $($i) of $($mailboxCount) mailboxes --- $($mailbox.UserPrincipalName)" -PercentComplete (($i * 100) / $mailboxCount)
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
            foreach ($emailAddress in $($mailbox.EmailAddresses.SmtpAddress.Where({ $_ -match $DomainName })))
            {
                $removeLegacyAddress = Remove-LegacyAddress -Mailbox $mailbox -LegacyAddress $emailAddress -ReportOnly $ReportOnly
                $output.Add([PSCustomObject]$removeLegacyAddress) | Out-Null
            }
            $j++
        }
    }
    else
    {
        Write-Verbose "Mailbox = $($mailbox.Identity)"
        foreach ($emailAddress in $($mailbox.EmailAddresses.SmtpAddress.Where({ $_ -match $DomainName })))
        {
            Write-Verbose "EmailAddress = $emailAddress"
            $removeLegacyAddress = Remove-LegacyAddress -Mailbox $mailbox -LegacyAddress $emailAddress -ReportOnly $ReportOnly
            $output.Add([PSCustomObject]$removeLegacyAddress) | Out-Null
        }
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
