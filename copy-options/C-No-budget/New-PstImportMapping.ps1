<#
.SYNOPSIS
    Builds the Microsoft 365 PST import mapping CSV from a folder of PST files and the toolkit's user mapping, and
    flags every PST it can't match to a destination mailbox.

.DESCRIPTION
    TL;DR: Microsoft's free PST import (Purview > Data lifecycle management > Import) needs a CSV that says which PST
    goes into which mailbox, and whether into the mailbox or its archive. One typo and that row fails, or lands in
    the wrong mailbox. This script writes that CSV for you:
      - Scans -PstFolder for PSTs named <alias>.pst, <upn>.pst, <alias>_archive.pst or <upn>_archive.pst
        (alias = the part of the source address before the @).
      - Matches each one to a Source in the mapping CSV (Source,Destination from 02b) and uses the Destination as the
        target mailbox. A name that matches no row, or more than one, is flagged and left out.
      - Writes PstImportMapping-<date>.csv with the exact columns the import service expects, plus a review CSV of
        every PST and what happened to it.
      - Prints the AzCopy command to upload the PSTs to the SAS URL the import page gives you.
    It reads local files only. It doesn't sign in to anything or change either tenant.

.PARAMETER PstFolder
    Local folder holding the PSTs (top level only, not subfolders).

.PARAMETER MappingFile
    Source,Destination CSV from 02b. Source is the old address, Destination the new mailbox.

.PARAMETER TargetRootFolder
    Where the PST's folders land in the mailbox. Default '/': merged into the mailbox's own folders (Inbox into
    Inbox). '/Imported' puts everything under a folder called Imported. Leaving the column blank in the CSV, which
    you can do with -TargetRootFolder '', makes the service create a folder called Imported.

.PARAMETER FilePath
    The folder in the Azure upload location the PSTs sit in. Default blank: the root of the SAS container. Set it
    (e.g. 'Batch1') if you upload into a subfolder; the AzCopy command printed at the end then targets it.

.PARAMETER OutputFolder
    Where the CSVs go. Default: the current folder.

.EXAMPLE
    .\New-PstImportMapping.ps1 -PstFolder C:\PST -MappingFile .\UserMapping.csv

.EXAMPLE
    .\New-PstImportMapping.ps1 -PstFolder C:\PST -MappingFile .\UserMapping.csv -TargetRootFolder /Imported -FilePath Batch1

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.0.0
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: Windows PowerShell 5.1 or PowerShell 7. No modules. AzCopy v10 for the upload itself.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$PstFolder,
    [Parameter(Mandatory)][string]$MappingFile,
    [AllowEmptyString()][string]$TargetRootFolder = '/',
    [AllowEmptyString()][string]$FilePath = '',
    [string]$OutputFolder = '.'
)

$ErrorActionPreference = 'Stop'
$stamp = Get-Date -Format 'yyyyMMdd-HHmm'
if (-not (Test-Path $PstFolder -PathType Container)) { throw "$PstFolder isn't a folder." }
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }
if ($TargetRootFolder -and -not $TargetRootFolder.StartsWith('/')) { $TargetRootFolder = '/' + $TargetRootFolder }
$FilePath = $FilePath.Trim('/', '\')

# ------------------------------------------------------------ mapping: index by full address and by alias
$mapRows = @(Import-Csv $MappingFile)
foreach ($c in 'Source', 'Destination') { if (-not ($mapRows | Get-Member -Name $c)) { throw "$MappingFile needs a '$c' column." } }
$byAddress = @{}
$byAlias = @{}
foreach ($r in $mapRows) {
    if (-not $r.Source -or -not $r.Destination) { continue }
    $src = $r.Source.Trim().ToLower()
    $dst = $r.Destination.Trim()
    $byAddress[$src] = $dst
    $alias = $src.Split('@')[0]
    if (-not $byAlias.ContainsKey($alias)) { $byAlias[$alias] = [System.Collections.Generic.List[string]]::new() }
    if (-not $byAlias[$alias].Contains($dst)) { $byAlias[$alias].Add($dst) }
}

# ------------------------------------------------------------ scan
$psts = @(Get-ChildItem -Path $PstFolder -Filter *.pst -File)
if ($psts.Count -eq 0) { throw "No .pst files in $PstFolder." }

$review = [System.Collections.Generic.List[object]]::new()
$mapping = [System.Collections.Generic.List[object]]::new()
foreach ($f in $psts) {
    $base = [IO.Path]::GetFileNameWithoutExtension($f.Name)
    $isArchive = $base -match '(?i)_archive$'
    $key = ($base -replace '(?i)_archive$', '').Trim().ToLower()
    $row = [ordered]@{ File = $f.Name; SizeGB = [math]::Round($f.Length / 1GB, 2); IsArchive = $isArchive.ToString().ToUpper(); Mailbox = ''; Status = ''; Note = '' }

    if ($key.Contains('@')) {
        if ($byAddress.ContainsKey($key)) { $row.Mailbox = $byAddress[$key] }
    }
    elseif ($byAlias.ContainsKey($key)) {
        if ($byAlias[$key].Count -eq 1) { $row.Mailbox = $byAlias[$key][0] }
        else { $row.Status = 'Ambiguous'; $row.Note = "Alias '$key' maps to $($byAlias[$key] -join ', '). Rename the PST to the full source address" }
    }

    if (-not $row.Status) {
        if ($row.Mailbox) { $row.Status = 'Matched' }
        else { $row.Status = 'Unmatched'; $row.Note = "No Source in $(Split-Path $MappingFile -Leaf) for '$key'. Rename the PST or add a mapping row" }
    }
    if ($f.Name -match '[^A-Za-z0-9._@\-]') { $row.Note = (($row.Note, 'Name has spaces or special characters: rename before upload') | Where-Object { $_ }) -join '. ' }
    if ($f.Length -gt 20GB) { $row.Note = (($row.Note, 'Over 20 GB: Microsoft recommends 20 GB or less per PST. Expect it to be slow') | Where-Object { $_ }) -join '. ' }
    $review.Add([pscustomobject]$row)

    if ($row.Status -eq 'Matched') {
        $mapping.Add([pscustomobject][ordered]@{
            Workload = 'Exchange'; FilePath = $FilePath; Name = $f.Name; Mailbox = $row.Mailbox; IsArchive = $row.IsArchive
            TargetRootFolder = $TargetRootFolder; ContentCodePage = ''; SPFileContainer = ''; SPManifestContainer = ''; SPSiteUrl = ''
        })
    }
}

# Two PSTs into the same mailbox (or archive) usually means alex.pst and alex@fabrikam.com.pst: both would import.
$mapping | Group-Object Mailbox, IsArchive | Where-Object Count -gt 1 | ForEach-Object {
    $files = $_.Group.Name
    foreach ($r in $review) { if ($files -contains $r.File) { $r.Note = (($r.Note, "Same target as $($files -join ' and '): both will import") | Where-Object { $_ }) -join '. ' } }
}

# ------------------------------------------------------------ write: exact header, quotes only where needed
function ConvertTo-CsvField([string]$v) { if ($v -match '[",\r\n]') { '"' + $v.Replace('"', '""') + '"' } else { $v } }
$cols = 'Workload', 'FilePath', 'Name', 'Mailbox', 'IsArchive', 'TargetRootFolder', 'ContentCodePage', 'SPFileContainer', 'SPManifestContainer', 'SPSiteUrl'
$lines = [System.Collections.Generic.List[string]]::new()
$lines.Add($cols -join ',')
foreach ($m in $mapping) { $lines.Add((($cols | ForEach-Object { ConvertTo-CsvField "$($m.$_)" }) -join ',')) }
$mapFile = Join-Path $OutputFolder "PstImportMapping-$stamp.csv"
[IO.File]::WriteAllLines((Join-Path (Resolve-Path $OutputFolder) "PstImportMapping-$stamp.csv"), $lines, [Text.UTF8Encoding]::new($false))
$revFile = Join-Path $OutputFolder "PstImport-Review-$stamp.csv"
$review | Export-Csv $revFile -NoTypeInformation

# ------------------------------------------------------------ summary and the upload command
$bad = @($review | Where-Object Status -ne 'Matched')
Write-Host ("{0} PSTs: {1} matched, {2} not. Mapping: {3}. Review: {4}" -f $psts.Count, $mapping.Count, $bad.Count, $mapFile, $revFile) -ForegroundColor Green
foreach ($b in $bad) { Write-Host ("  {0,-10} {1}: {2}" -f $b.Status, $b.File, $b.Note) -ForegroundColor Yellow }
foreach ($d in @($review | Where-Object { $_.Status -eq 'Matched' -and $_.Note })) { Write-Host ("  Check      {0}: {1}" -f $d.File, $d.Note) -ForegroundColor Yellow }
$total = [math]::Round(($psts | Measure-Object Length -Sum).Sum / 1GB, 1)
$src = (Resolve-Path $PstFolder).Path.TrimEnd('\', '/')
$dest = if ($FilePath) { "<SAS URL with /$FilePath inserted before the ?>" } else { '<SAS URL>' }
Write-Host ''
Write-Host "Upload ($total GB) with AzCopy v10:" -ForegroundColor Cyan
Write-Host ("  azcopy copy ""{0}{1}*"" ""{2}"" --recursive" -f $src, [IO.Path]::DirectorySeparatorChar, $dest)
Write-Host 'Tip: the SAS URL comes from Purview > Data lifecycle management > Import > New import job > Upload your data > Show network upload SAS URL.' -ForegroundColor Yellow
Write-Host "Trap: copy the folder's contents (the trailing *), not the folder. Without it AzCopy adds the folder name to the path and every row's FilePath is wrong." -ForegroundColor Yellow
Write-Host 'Trap: the SAS URL is a key to the upload location. Keep it out of tickets, chat and this repo.' -ForegroundColor Yellow
