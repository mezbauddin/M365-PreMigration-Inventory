<#
.SYNOPSIS
    Copies OneDrives or document libraries between tenants with nothing but Microsoft Graph: download from the source
    to a local folder, then upload into the destination, keeping folders and file dates.

.DESCRIPTION
    TL;DR: Two passes, because a Graph session can only be in one tenant at a time.
      -Download   Signed in to the source. Pulls each source site's document library into <LocalPath>\<name>.
      -Upload     Signed in to the destination. Pushes every <LocalPath>\<name> folder into <library>\<name> on a
                  destination site.
    Both passes skip files they've already done (same name and size), so you can stop and re-run. Large files go up in
    10 MB chunks. File dates are kept. What isn't kept: version history, created-by and modified-by (they become you),
    sharing links and unique permissions. Fine for leavers' OneDrives and small libraries; use a migration tool for
    anything people work in every day.

.PARAMETER Download
    Source pass.

.PARAMETER Upload
    Destination pass.

.PARAMETER TenantId
    Tenant ID (GUID) for this pass: the source for -Download, the destination for -Upload.

.PARAMETER SourceSiteUrl
    -Download: OneDrive or site URLs, e.g. https://fabrikam-my.sharepoint.com/personal/alex_fabrikam_com. Each one's
    folder is named after the part before the first underscore for OneDrives, or the last URL segment for sites.

.PARAMETER SourceLibrary
    -Download: library to copy. Default 'Documents' (the default library of the site).

.PARAMETER DestinationSiteUrl
    -Upload: the destination site, e.g. https://contoso.sharepoint.com/sites/Archive.

.PARAMETER DestinationLibrary
    -Upload: library name on that site. Each local folder lands in a folder of the same name inside it.

.PARAMETER LocalPath
    Working folder. Keep it short (e.g. C:\T2T) to stay clear of Windows path limits. Default .\DriveCopy.

.PARAMETER Name
    Only these local folder names. Default: every folder under LocalPath (or every row of -CsvPath).

.PARAMETER CsvPath
    Optional. The copy list from ..\New-DriveCopyList.ps1 (the same CSV the ShareGate script reads). -Download
    pulls every Source into <LocalPath>\<DestinationFolder>; -Upload pushes each one to its own DestinationSite,
    DestinationLibrary and DestinationFolder (a blank folder means the library root, for OneDrive to OneDrive).
    Rows with Status 'Skip' are left out. Replaces -SourceSiteUrl and -DestinationSiteUrl.

.EXAMPLE
    .\Copy-DriveViaGraph.ps1 -Download -TenantId <source-guid> -LocalPath C:\T2T `
        -SourceSiteUrl https://fabrikam-my.sharepoint.com/personal/alex_fabrikam_com

.EXAMPLE
    .\Copy-DriveViaGraph.ps1 -Upload -TenantId <dest-guid> -LocalPath C:\T2T `
        -DestinationSiteUrl https://contoso.sharepoint.com/sites/Archive -DestinationLibrary 'Former Staff OneDrive'

.EXAMPLE
    .\Copy-DriveViaGraph.ps1 -Download -TenantId <source-guid> -LocalPath C:\T2T -CsvPath ..\DriveCopyList-20261110-2100.csv
    .\Copy-DriveViaGraph.ps1 -Upload -TenantId <dest-guid> -LocalPath C:\T2T -CsvPath ..\DriveCopyList-20261110-2100.csv

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.1.0
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: Microsoft.Graph.Authentication. You need access to every source drive (make yourself site collection
              admin of each OneDrive) and write access to the destination library. Archived OneDrives can't be read.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(ParameterSetName = 'Download', Mandatory)][switch]$Download,
    [Parameter(ParameterSetName = 'Upload', Mandatory)][switch]$Upload,
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F-]{36}$')][string]$TenantId,
    [string]$LocalPath = '.\DriveCopy',

    [string]$CsvPath,

    [Parameter(ParameterSetName = 'Download')][string[]]$SourceSiteUrl,
    [Parameter(ParameterSetName = 'Download')][string]$SourceLibrary = 'Documents',

    [Parameter(ParameterSetName = 'Upload')][string]$DestinationSiteUrl,
    [Parameter(ParameterSetName = 'Upload')][string]$DestinationLibrary,
    [Parameter(ParameterSetName = 'Upload')][string[]]$Name
)

$ErrorActionPreference = 'Stop'
$G = 'https://graph.microsoft.com/v1.0'

# ------------------------------------------------------------ what to copy: parameters, or the CSV from New-DriveCopyList
function Get-LocalName([string]$url, [string]$folder, [bool]$fromCsv) {
    if ($folder) { return $folder }
    $seg = ([uri]$url.TrimEnd('/')).Segments[-1].TrimEnd('/')
    # From a CSV, a OneDrive-to-OneDrive row keeps the whole URL segment so alex_smith and alex_jones can't collide.
    if ($fromCsv) { return $seg }
    if ($url -match '/personal/') { $seg.Split('_')[0] } else { $seg }
}
$jobs = [System.Collections.Generic.List[object]]::new()
if ($CsvPath) {
    $csvRows = @(Import-Csv $CsvPath)
    foreach ($c in 'Source', 'DestinationSite') { if (-not ($csvRows | Get-Member -Name $c)) { throw "$CsvPath needs a '$c' column." } }
    foreach ($r in $csvRows) {
        if ($r.Status -eq 'Skip' -or -not $r.Source -or -not $r.DestinationSite) { Write-Host "Skipping $(if ($r.Owner) { $r.Owner } else { $r.Source }): $($r.Reason)" -ForegroundColor Yellow; continue }
        $jobs.Add([pscustomobject]@{
            Source = $r.Source; SourceLibrary = $(if ($r.SourceLibrary) { $r.SourceLibrary } else { $SourceLibrary })
            Local = Get-LocalName $r.Source $r.DestinationFolder $true
            Site = $r.DestinationSite; Library = $(if ($r.DestinationLibrary) { $r.DestinationLibrary } else { 'Documents' })
            Remote = "$($r.DestinationFolder)"
        })
    }
}
elseif ($Download) {
    if (-not $SourceSiteUrl) { throw 'Pass -SourceSiteUrl or -CsvPath.' }
    foreach ($url in $SourceSiteUrl) { $jobs.Add([pscustomobject]@{ Source = $url; SourceLibrary = $SourceLibrary; Local = Get-LocalName $url '' $false }) }
}
else {
    if (-not $DestinationSiteUrl -or -not $DestinationLibrary) { throw 'Pass -DestinationSiteUrl and -DestinationLibrary, or -CsvPath.' }
}
if ($Name) { $jobs = [System.Collections.Generic.List[object]]@($jobs | Where-Object { $Name -contains $_.Local }) }
$scopes = if ($Download) { 'Sites.Read.All', 'Files.Read.All' } else { 'Sites.ReadWrite.All', 'Files.ReadWrite.All' }

Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
Connect-MgGraph -TenantId $TenantId -Scopes $scopes -NoWelcome
$ctx = Get-MgContext
if ($ctx.TenantId -ne $TenantId) { throw "Graph signed in to $($ctx.TenantId), not $TenantId." }
Write-Host "Signed in to $TenantId as $($ctx.Account)" -ForegroundColor Cyan

New-Item -ItemType Directory -Path $LocalPath -Force -WhatIf:$false | Out-Null
$LocalPath = (Resolve-Path $LocalPath).Path
$log = [System.Collections.Generic.List[object]]::new()
function Write-Log($Who, $Path, $Status, $Detail) {
    $log.Add([pscustomobject]@{ Name = $Who; Path = $Path; Status = $Status; Detail = $Detail })
    if ($Status -in 'Problem', 'Skipped') { Write-Host "  [$Status] $Who $Path $Detail" -ForegroundColor $(if ($Status -eq 'Problem') { 'Red' } else { 'Yellow' }) }
}

function Get-SiteId([string]$url) {
    $u = [uri]$url.TrimEnd('/')
    (Invoke-MgGraphRequest GET "$G/sites/$($u.Host):$($u.AbsolutePath)").id
}
function Get-DriveId([string]$siteId, [string]$library) {
    $drives = (Invoke-MgGraphRequest GET "$G/sites/$siteId/drives").value
    $d = $drives | Where-Object { $_.name -eq $library } | Select-Object -First 1
    if (-not $d -and $library -eq 'Documents') { $d = Invoke-MgGraphRequest GET "$G/sites/$siteId/drive" }
    if (-not $d) { throw "No library called '$library'. Found: $(($drives.name) -join ', ')" }
    $d.id
}
function Get-EncodedPath([string]$rel) { ($rel -split '[\\/]' | Where-Object { $_ } | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/' }

# ============================================================ DOWNLOAD
if ($Download) {
    foreach ($job in $jobs) {
        $url = $job.Source
        $who = $job.Local
        Write-Host "== $who" -ForegroundColor Cyan
        try { $drive = Get-DriveId (Get-SiteId $url) $job.SourceLibrary }
        catch { Write-Log $who $url 'Problem' "Can't open: $($_.Exception.Message). Archived OneDrive, or you're not an admin of it?"; continue }

        $root = Join-Path $LocalPath $who
        New-Item -ItemType Directory -Path $root -Force -WhatIf:$false | Out-Null
        $files = 0; $bytes = 0
        $stack = [System.Collections.Generic.Stack[object]]::new()
        $stack.Push(@('root', ''))
        while ($stack.Count) {
            $itemId, $rel = $stack.Pop()
            $uri = "$G/drives/$drive/items/$itemId/children?`$top=200&`$select=id,name,size,file,folder,package,fileSystemInfo"
            do {
                $page = Invoke-MgGraphRequest GET $uri
                foreach ($i in $page.value) {
                    $childRel = if ($rel) { "$rel\$($i.name)" } else { $i.name }
                    if ($i.folder) { New-Item -ItemType Directory -Path (Join-Path $root $childRel) -Force -WhatIf:$false | Out-Null; $stack.Push(@($i.id, $childRel)); continue }
                    if ($i.package) { Write-Log $who $childRel 'Skipped' 'OneNote notebook or package: export it separately'; continue }
                    $out = Join-Path $root $childRel
                    if ((Test-Path -LiteralPath $out) -and (Get-Item -LiteralPath $out).Length -eq [int64]$i.size) { $files++; $bytes += $i.size; continue }
                    try {
                        Invoke-MgGraphRequest GET "$G/drives/$drive/items/$($i.id)/content" -OutputFilePath $out
                        $f = Get-Item -LiteralPath $out
                        if ($i.fileSystemInfo.lastModifiedDateTime) { $f.LastWriteTimeUtc = [datetime]$i.fileSystemInfo.lastModifiedDateTime }
                        if ($i.fileSystemInfo.createdDateTime) { $f.CreationTimeUtc = [datetime]$i.fileSystemInfo.createdDateTime }
                        $files++; $bytes += $i.size
                    }
                    catch { Write-Log $who $childRel 'Problem' $_.Exception.Message }
                }
                $uri = $page.'@odata.nextLink'
            } while ($uri)
        }
        Write-Log $who $root 'Done' "$files files, $([math]::Round($bytes / 1MB, 1)) MB"
        Write-Host ("   {0} files, {1} MB" -f $files, [math]::Round($bytes / 1MB, 1)) -ForegroundColor Green
    }
}

# ============================================================ UPLOAD
if ($Upload) {
    # Without a CSV: every local folder goes into <library>\<folder name> on one site, as before.
    if (-not $CsvPath) {
        foreach ($d in Get-ChildItem -LiteralPath $LocalPath -Directory | Where-Object { -not $Name -or $Name -contains $_.Name }) {
            $jobs.Add([pscustomobject]@{ Local = $d.Name; Site = $DestinationSiteUrl; Library = $DestinationLibrary; Remote = $d.Name })
        }
    }
    $made = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $driveCache = @{}

    function Confirm-RemoteFolder([string]$rel) {
        $parts = $rel -split '[\\/]' | Where-Object { $_ }
        $path = ''
        foreach ($p in $parts) {
            $parent = $path
            $path = if ($path) { "$path/$p" } else { $p }
            if ($made.Contains("$drive|$path")) { continue }
            if ($WhatIfPreference) { [void]$made.Add("$drive|$path"); continue }   # -WhatIf: don't create folders either
            $uri = if ($parent) { "$G/drives/$drive/root:/$(Get-EncodedPath $parent):/children" } else { "$G/drives/$drive/root/children" }
            $body = @{ name = $p; folder = @{}; '@microsoft.graph.conflictBehavior' = 'fail' } | ConvertTo-Json
            try { Invoke-MgGraphRequest POST $uri -Body $body -ContentType 'application/json' | Out-Null }
            catch { if ("$($_.Exception.Message)$($_.ErrorDetails.Message)" -notmatch 'nameAlreadyExists|409|Conflict') { throw } }
            [void]$made.Add("$drive|$path")
        }
    }
    function Join-Remote([string]$a, [string]$b) { if ($a) { if ($b) { Join-Path $a $b } else { $a } } else { $b } }

    foreach ($job in $jobs) {
        $who = $job.Local
        $folder = Get-Item -LiteralPath (Join-Path $LocalPath $who) -ErrorAction SilentlyContinue
        if (-not $folder) { Write-Log $who (Join-Path $LocalPath $who) 'Problem' 'Local folder not found. Run -Download first'; continue }
        $key = "$($job.Site)|$($job.Library)"
        try {
            if (-not $driveCache.ContainsKey($key)) { $driveCache[$key] = Get-DriveId (Get-SiteId $job.Site) $job.Library }
            $drive = $driveCache[$key]
        }
        catch { Write-Log $who $job.Site 'Problem' "Can't open the destination library: $($_.Exception.Message)"; continue }
        Write-Host "== $who -> $($job.Site) [$($job.Library)$(if ($job.Remote) { "/$($job.Remote)" })]" -ForegroundColor Cyan
        $files = 0; $bytes = 0; $skipped = 0
        if ($job.Remote) { Confirm-RemoteFolder $job.Remote }
        foreach ($dir in Get-ChildItem -LiteralPath $folder.FullName -Directory -Recurse) {
            Confirm-RemoteFolder (Join-Remote $job.Remote $dir.FullName.Substring($folder.FullName.Length + 1))
        }
        foreach ($f in Get-ChildItem -LiteralPath $folder.FullName -File -Recurse) {
            $rel = Join-Remote $job.Remote $f.FullName.Substring($folder.FullName.Length + 1)
            $enc = Get-EncodedPath $rel
            try {
                $existing = $null
                try { $existing = Invoke-MgGraphRequest GET "$G/drives/$drive/root:/${enc}?`$select=size" } catch {}
                if ($existing -and [int64]$existing.size -eq $f.Length) { $skipped++; continue }
                if (-not $PSCmdlet.ShouldProcess($rel, 'Upload')) { continue }

                $times = @{ lastModifiedDateTime = $f.LastWriteTimeUtc.ToString('o'); createdDateTime = $f.CreationTimeUtc.ToString('o') }
                if ($f.Length -eq 0) {
                    Invoke-MgGraphRequest PUT "$G/drives/$drive/root:/${enc}:/content" -Body ([byte[]]@()) -ContentType 'application/octet-stream' | Out-Null
                }
                else {
                    $sessionBody = @{ item = @{ '@microsoft.graph.conflictBehavior' = 'replace'; fileSystemInfo = $times } } | ConvertTo-Json -Depth 4
                    $session = Invoke-MgGraphRequest POST "$G/drives/$drive/root:/${enc}:/createUploadSession" -Body $sessionBody -ContentType 'application/json'
                    $chunk = 320KB * 32   # 10 MB, a multiple of 320 KiB as Graph requires
                    $stream = [IO.File]::OpenRead($f.FullName)
                    try {
                        $buffer = New-Object byte[] $chunk
                        $pos = 0
                        while ($pos -lt $f.Length) {
                            $read = $stream.Read($buffer, 0, $chunk)
                            $body = if ($read -eq $chunk) { $buffer } else { $buffer[0..($read - 1)] }
                            $range = "bytes $pos-$($pos + $read - 1)/$($f.Length)"
                            # The upload URL is pre-authorised: no Graph token on these calls.
                            Invoke-RestMethod -Method Put -Uri $session.uploadUrl -Body ([byte[]]$body) -Headers @{ 'Content-Range' = $range } -ContentType 'application/octet-stream' | Out-Null
                            $pos += $read
                        }
                    }
                    finally { $stream.Dispose() }
                }
                $files++; $bytes += $f.Length
            }
            catch { Write-Log $who $rel 'Problem' $(if ($_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }) }
        }
        Write-Log $who $who 'Done' "$files uploaded, $skipped already there, $([math]::Round($bytes / 1MB, 1)) MB"
        Write-Host ("   {0} uploaded, {1} already there, {2} MB" -f $files, $skipped, [math]::Round($bytes / 1MB, 1)) -ForegroundColor Green
    }
}

$file = Join-Path $LocalPath "DriveCopy-$(if ($Download) { 'Download' } else { 'Upload' })-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
$log | Export-Csv $file -NoTypeInformation -WhatIf:$false
$problems = @($log | Where-Object Status -eq 'Problem').Count
Write-Host ''
Write-Host "Report: $file" -ForegroundColor Cyan
if ($problems) { Write-Host "$problems problem(s). Fix and re-run: finished files are skipped." -ForegroundColor Red }
else { Write-Host 'No problems.' -ForegroundColor Green }
