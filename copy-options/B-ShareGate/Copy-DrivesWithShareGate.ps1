<#
.SYNOPSIS
    Copies many OneDrives (or libraries) between tenants in one run with ShareGate's PowerShell module, either into
    the matching OneDrive or into a folder of a library, with your user mapping applied.

.DESCRIPTION
    TL;DR: The ShareGate app copies OneDrives one at a time. This script reads a CSV of source and destination pairs,
    signs in to each tenant once, and runs Copy-Content for every row, with optional incremental mode, a .sgum user
    mapping and a pre-check (-WhatIf). Each row can land in a whole library (OneDrive to OneDrive) or in a folder of a
    library (for example leavers' OneDrives into an archive library). Writes a report of successes, warnings and
    errors per row.

    CSV columns (header row required):
      Source              Source OneDrive or site URL
      DestinationSite     Destination OneDrive or site URL
      DestinationLibrary  Optional. Library on the destination. Default 'Documents'
      DestinationFolder   Optional. Folder inside that library. It must already exist
      SourceLibrary       Optional. Default 'Documents'

.PARAMETER CsvPath
    The CSV described above.

.PARAMETER SourceAdminUrl
    The source tenant's SharePoint admin URL. Used for one browser sign-in that every source row reuses.

.PARAMETER DestinationAdminUrl
    The destination tenant's SharePoint admin URL. One browser sign-in for every destination row.

.PARAMETER MappingFile
    Optional ShareGate user and group mapping file (.sgum), so authors and permissions land on the right people.

.PARAMETER Incremental
    Only copy what's new or changed since the last run (ShareGate incremental update).

.PARAMETER WhatIf
    Run ShareGate's pre-check for every row instead of copying.

.EXAMPLE
    .\Copy-DrivesWithShareGate.ps1 -CsvPath .\drives.csv -SourceAdminUrl https://fabrikam-admin.sharepoint.com `
        -DestinationAdminUrl https://contoso-admin.sharepoint.com -MappingFile .\UserMapping.sgum -WhatIf

.EXAMPLE
    .\Copy-DrivesWithShareGate.ps1 -CsvPath .\drives.csv -SourceAdminUrl https://fabrikam-admin.sharepoint.com `
        -DestinationAdminUrl https://contoso-admin.sharepoint.com -MappingFile .\UserMapping.sgum -Incremental

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.0.1
    Date    : 2026-11-09
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: ShareGate Migrate installed on this PC, with a Pro or Enterprise subscription (PowerShell isn't in
              Essentials). Windows PowerShell 5.1 only (the module won't load in PowerShell 7). You need site collection admin on every source OneDrive, and the destination
              OneDrives must already exist (part 04 of this toolkit). Not sponsored by ShareGate.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$CsvPath,
    [Parameter(Mandatory)][string]$SourceAdminUrl,
    [Parameter(Mandatory)][string]$DestinationAdminUrl,
    [string]$MappingFile,
    [switch]$Incremental,
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSEdition -eq 'Core') {
    throw 'ShareGate''s module only loads in Windows PowerShell 5.1, not PowerShell 7. Run: powershell.exe -ExecutionPolicy Bypass -File .\Copy-DrivesWithShareGate.ps1 <same parameters>'
}
if (-not (Get-Module -ListAvailable -Name Sharegate)) {
    throw 'The ShareGate PowerShell module isn''t available. Run this from the ShareGate PowerShell shortcut, or install ShareGate Migrate on this PC.'
}
Import-Module Sharegate

$rows = @(Import-Csv $CsvPath)
foreach ($c in 'Source', 'DestinationSite') {
    if (-not ($rows | Get-Member -Name $c)) { throw "$CsvPath needs a '$c' column." }
}
Write-Host "$($rows.Count) row(s) to copy." -ForegroundColor Cyan

# One browser sign-in per tenant. Every row reuses it.
Write-Host 'Sign in to the SOURCE tenant.' -ForegroundColor Magenta
$srcConn = Connect-Site -Url $SourceAdminUrl -Browser
Write-Host 'Sign in to the DESTINATION tenant.' -ForegroundColor Magenta
$dstConn = Connect-Site -Url $DestinationAdminUrl -Browser

$mappings = $null
if ($MappingFile) {
    $mappings = New-MappingSettings
    $mappings = Import-UserAndGroupMapping -MappingSettings $mappings -Path (Resolve-Path $MappingFile).Path
    Write-Host "User mapping loaded from $MappingFile" -ForegroundColor Cyan
}
$copySettings = if ($Incremental) { New-CopySettings -OnContentItemExists IncrementalUpdate } else { $null }

$report = foreach ($row in $rows) {
    $name = ([uri]$row.Source.TrimEnd('/')).Segments[-1].TrimEnd('/')
    $srcLibName = if ($row.SourceLibrary) { $row.SourceLibrary } else { 'Documents' }
    $dstLibName = if ($row.DestinationLibrary) { $row.DestinationLibrary } else { 'Documents' }
    $target = if ($row.DestinationFolder) { "$dstLibName/$($row.DestinationFolder)" } else { $dstLibName }
    Write-Host "== $name -> $($row.DestinationSite) [$target]" -ForegroundColor Cyan
    try {
        $srcSite = Connect-Site -Url $row.Source -UseCredentialsFrom $srcConn
        $dstSite = Connect-Site -Url $row.DestinationSite -UseCredentialsFrom $dstConn
        $srcList = Get-List -Site $srcSite -Name $srcLibName
        $dstList = Get-List -Site $dstSite -Name $dstLibName

        $p = @{ SourceList = $srcList; DestinationList = $dstList; TaskName = "T2T $name" }
        if ($row.DestinationFolder) { $p.DestinationFolder = $row.DestinationFolder }
        if ($mappings) { $p.MappingSettings = $mappings }
        if ($copySettings) { $p.CopySettings = $copySettings }
        if ($WhatIf) { $p.WhatIf = $true }

        $r = Copy-Content @p
        Write-Host ("   {0}: {1} successes, {2} warnings, {3} errors" -f $r.Result, $r.Successes, $r.Warnings, $r.Errors) -ForegroundColor $(if ($r.Errors) { 'Yellow' } else { 'Green' })
        [pscustomobject]@{ Source = $row.Source; Destination = "$($row.DestinationSite) [$target]"; Result = $r.Result
                           Successes = $r.Successes; Warnings = $r.Warnings; Errors = $r.Errors; SessionId = $r.SessionId; Note = '' }
    }
    catch {
        Write-Host "   Problem: $($_.Exception.Message)" -ForegroundColor Red
        [pscustomobject]@{ Source = $row.Source; Destination = "$($row.DestinationSite) [$target]"; Result = 'Failed'
                           Successes = 0; Warnings = 0; Errors = 0; SessionId = ''; Note = $_.Exception.Message }
    }
}

$file = ".\DriveCopy-ShareGate-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
$report | Export-Csv $file -NoTypeInformation
Write-Host ''
$report | Format-Table Source, Result, Successes, Warnings, Errors -AutoSize | Out-Host
Write-Host "Report: $file. Item-level detail is in ShareGate's Tasks, under each 'T2T <name>' task." -ForegroundColor Cyan
