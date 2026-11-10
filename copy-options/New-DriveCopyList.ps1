<#
.SYNOPSIS
    Builds the OneDrive copy list (CSV) that Copy-DrivesWithShareGate.ps1 (option B) and Copy-DriveViaGraph.ps1
    (option C) read: leavers into one folder each in an archive library, current staff into their new OneDrive.

.DESCRIPTION
    TL;DR: Writing the copy CSV by hand means typing dozens of OneDrive URLs, guessing who has left, and creating a
    folder per person in the archive library. This does it for you. It reads every OneDrive in the source, picks the
    people you name (-Users) or every leaver (-Leavers: OneDrive owners who are unlicensed, blocked from signing in,
    or already deleted), and writes one row per OneDrive:

      Source              https://fabrikam-my.sharepoint.com/personal/kim_fabrikam_com
      DestinationSite     the archive site, or the person's new OneDrive (with -MappingFile)
      DestinationLibrary  the archive library, or Documents for a OneDrive
      DestinationFolder   one folder per person in the archive library, blank for a OneDrive
      SourceLibrary       Documents
      Status, Reason      Copy, or Skip with the reason (archived, locked, no destination, not found)
      Owner, DisplayName, StorageGB, LastModified   for your review

    OneDrives in Microsoft 365 Archive (ArchiveStatus anything but NotArchived) can't be read by any copy tool until
    they are reactivated, so they are marked Skip. Both copy scripts skip rows marked Skip.
    -CreateFolders creates the per-person folders in the destination library through Microsoft Graph, so the
    ShareGate script (which needs them to exist) can run straight after.

.PARAMETER SourceAdminUrl
    The source SharePoint admin URL, e.g. https://fabrikam-admin.sharepoint.com

.PARAMETER Users
    Sign-in names (kim@fabrikam.com) or user names (kim) whose OneDrives to include. A user name must match exactly
    one OneDrive.

.PARAMETER Leavers
    Add every OneDrive whose owner is unlicensed, blocked from signing in, or no longer exists in the source. Reads
    users through Microsoft Graph (User.Read.All) in the source tenant.

.PARAMETER DestinationSite
    The destination site for leavers, e.g. https://contoso.sharepoint.com/sites/Archive

.PARAMETER DestinationLibrary
    The library on DestinationSite, e.g. 'Former Staff OneDrive'. Default 'Documents'.

.PARAMETER MappingFile
    Optional. UserMapping.csv from 02b (Source,Destination). People in it get their new OneDrive as the destination
    instead of a folder: https://<destination>-my.sharepoint.com/personal/<new sign-in>. Rebuild it after cutover
    first, because sign-in names (and OneDrive URLs) change.

.PARAMETER DestinationAdminUrl
    Optional destination SharePoint admin URL. Only needed with -MappingFile when there's no -DestinationSite to
    work out the destination's OneDrive host from.

.PARAMETER FolderName
    How to name each person's folder: UserName (default, the part of the sign-in before @) or DisplayName.

.PARAMETER CreateFolders
    Create each DestinationFolder in DestinationLibrary through Microsoft Graph (Sites.ReadWrite.All) in the
    destination tenant. Existing folders are left alone. Supports -WhatIf.

.PARAMETER SourceTenantId
    Optional source tenant ID or domain for the Graph sign-in with -Leavers. Default: <name>.onmicrosoft.com worked
    out from -SourceAdminUrl.

.PARAMETER DestinationTenantId
    Optional destination tenant ID or domain for -CreateFolders. Default: worked out from -DestinationSite.

.PARAMETER AdminUpn
    Optional. The sign-in you expect for Graph. If Graph is signed in as anyone else it disconnects first, and stops
    if the new sign-in doesn't match. Give the source admin with -Leavers, the destination admin with -CreateFolders
    (run the two separately if they differ).

.PARAMETER OutFile
    Where to write the CSV. Default .\DriveCopyList-<date>.csv

.EXAMPLE
    .\New-DriveCopyList.ps1 -SourceAdminUrl https://fabrikam-admin.sharepoint.com -Leavers `
        -DestinationSite https://contoso.sharepoint.com/sites/Archive -DestinationLibrary 'Former Staff OneDrive'

.EXAMPLE
    .\New-DriveCopyList.ps1 -SourceAdminUrl https://fabrikam-admin.sharepoint.com -Users kim, pat@fabrikam.com `
        -DestinationSite https://contoso.sharepoint.com/sites/Archive -DestinationLibrary 'Former Staff OneDrive' -CreateFolders -WhatIf

.EXAMPLE
    .\New-DriveCopyList.ps1 -SourceAdminUrl https://fabrikam-admin.sharepoint.com -Users alex, sam `
        -MappingFile .\UserMapping.csv -DestinationAdminUrl https://contoso-admin.sharepoint.com

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.0.0
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: Microsoft.Online.SharePoint.PowerShell and SharePoint Administrator in the source. For -Leavers and
              -CreateFolders: Microsoft.Graph.Authentication (all Microsoft.Graph modules at one version; see
              tools/Test-ModuleHealth.ps1). Run it in PowerShell 7; then run the ShareGate script in 5.1.
    The CSV lists people and their OneDrive URLs. Keep it out of source control and shared drives.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][string]$SourceAdminUrl,
    [string[]]$Users = @(),
    [switch]$Leavers,
    [string]$DestinationSite,
    [string]$DestinationLibrary = 'Documents',
    [string]$MappingFile,
    [string]$DestinationAdminUrl,
    [ValidateSet('UserName', 'DisplayName')][string]$FolderName = 'UserName',
    [switch]$CreateFolders,
    [string]$SourceTenantId,
    [string]$DestinationTenantId,
    [string]$AdminUpn,
    [string]$OutFile = ".\DriveCopyList-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
)

$ErrorActionPreference = 'Stop'
if (-not $Users -and -not $Leavers) { throw 'Pass -Users, -Leavers, or both.' }
if ($SourceAdminUrl -notmatch '^https://') { $SourceAdminUrl = "https://$SourceAdminUrl" }
if ($DestinationSite -and $DestinationSite -notmatch '^https://') { $DestinationSite = "https://$DestinationSite" }
if ($CreateFolders -and -not $DestinationSite) { throw '-CreateFolders needs -DestinationSite and -DestinationLibrary.' }
if ($MappingFile -and -not (Test-Path $MappingFile)) { throw "$MappingFile not found." }

function Get-TenantName([string]$url) {
    # https://fabrikam-admin.sharepoint.com, https://fabrikam.sharepoint.com/sites/x, https://fabrikam-my.sharepoint.com -> fabrikam
    if (-not $url) { return $null }
    ([uri]$url).Host.Split('.')[0] -replace '-(admin|my)$', ''
}
function Get-PersonalUrl([string]$tenant, [string]$upn) { "https://$tenant-my.sharepoint.com/personal/$($upn.ToLower() -replace '[@.]', '_')" }
function Get-SafeFolder([string]$n) { (($n -replace '["*:<>?/\\|#%]', '') -replace '\s+', ' ').Trim().TrimEnd('.') }
function Connect-GraphAs([string]$tenant, [string[]]$scopes) {
    # Always start fresh: a Graph session left open from the other tenant is how copies land in the wrong place.
    if (Get-MgContext) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    Connect-MgGraph -TenantId $tenant -Scopes $scopes -NoWelcome
    $ctx = Get-MgContext
    if ($AdminUpn -and $ctx.Account -ne $AdminUpn) { throw "Graph signed in as $($ctx.Account), not $AdminUpn. Run Disconnect-MgGraph and pick the right account." }
    Write-Host "Signed in as $($ctx.Account) to tenant $($ctx.TenantId)" -ForegroundColor Cyan
}

$srcTenant = Get-TenantName $SourceAdminUrl
$dstTenant = if ($DestinationSite) { Get-TenantName $DestinationSite } else { Get-TenantName $DestinationAdminUrl }
if ($MappingFile -and -not $dstTenant) { throw '-MappingFile needs -DestinationSite or -DestinationAdminUrl, to work out the destination OneDrive host.' }

# ------------------------------------------------------------ mapping
$map = @{}
if ($MappingFile) {
    foreach ($r in Import-Csv $MappingFile) {
        if (-not $r.Source -or -not $r.Destination) { continue }
        $map[$r.Source.Trim().ToLower()] = $r.Destination.Trim().ToLower()
    }
    Write-Host "$($map.Count) mapping row(s) from $MappingFile" -ForegroundColor Cyan
}

# ------------------------------------------------------------ source OneDrives
if (-not (Get-Command Connect-SPOService -ErrorAction SilentlyContinue)) {
    try { Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -ErrorAction Stop }
    catch {
        if ($PSVersionTable.PSEdition -eq 'Core') { Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -DisableNameChecking -WarningAction SilentlyContinue }
        else { throw 'Microsoft.Online.SharePoint.PowerShell is not installed. Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser' }
    }
}
Write-Host 'Sign in to the SOURCE SharePoint admin centre.' -ForegroundColor Magenta
Connect-SPOService -Url $SourceAdminUrl
$drives = @(Get-SPOSite -IncludePersonalSite $true -Limit All -Filter "Url -like '-my.sharepoint.com/personal/'")
Write-Host "$($drives.Count) OneDrive(s) in the source." -ForegroundColor Cyan
$byOwner = @{}; $byUrl = @{}
foreach ($d in $drives) {
    if ($d.Owner) { $byOwner["$($d.Owner)".ToLower()] = $d }
    $byUrl["$($d.Url)".TrimEnd('/').ToLower()] = $d
}

# ------------------------------------------------------------ who
$wanted = [ordered]@{}   # key: OneDrive URL -> @{ Upn; Why; Site }
$notFound = [System.Collections.Generic.List[object]]::new()
foreach ($u in $Users) {
    $u = $u.Trim().ToLower()
    $hit = $null
    if ($u -match '@') {
        $hit = $byOwner[$u]
        if (-not $hit) { $hit = $byUrl[(Get-PersonalUrl $srcTenant $u)] }
    }
    else {
        $m = @($drives | Where-Object { "$($_.Owner)".ToLower() -like "$u@*" })
        if (-not $m) { $m = @($drives | Where-Object { ([uri]$_.Url).Segments[-1].TrimEnd('/').ToLower() -like "$($u -replace '\.', '_')_*" }) }
        if ($m.Count -eq 1) { $hit = $m[0] }
        elseif ($m.Count -gt 1) { Write-Warning "'$u' matches $($m.Count) OneDrives ($(($m.Url) -join ', ')). Use the full sign-in name."; $notFound.Add(@($u, "'$u' matches $($m.Count) OneDrives. Use the full sign-in name")); continue }
    }
    if ($hit) { $wanted["$($hit.Url)".TrimEnd('/')] = @{ Upn = "$($hit.Owner)".ToLower(); Why = 'Named'; Site = $hit } }
    else { $notFound.Add(@($u, "No OneDrive found for '$u'. Never signed in, already deleted, or a typo")) }
}

if ($Leavers) {
    $graphTenant = if ($SourceTenantId) { $SourceTenantId } else { "$srcTenant.onmicrosoft.com" }
    Connect-GraphAs $graphTenant @('User.Read.All')
    $people = @{}
    $uri = 'https://graph.microsoft.com/v1.0/users?$select=id,userPrincipalName,displayName,accountEnabled,assignedLicenses&$top=999'
    do {
        $page = Invoke-MgGraphRequest GET $uri
        foreach ($p in $page.value) { $people["$($p.userPrincipalName)".ToLower()] = $p }
        $uri = $page.'@odata.nextLink'
    } while ($uri)
    Write-Host "$($people.Count) user(s) read from the source directory." -ForegroundColor Cyan
    foreach ($d in $drives) {
        $owner = "$($d.Owner)".ToLower()
        $p = $people[$owner]
        $why = if (-not $owner) { $null }
               elseif (-not $p) { 'Leaver: owner no longer in the directory' }
               elseif (-not $p.accountEnabled) { 'Leaver: sign-in blocked' }
               elseif (-not @($p.assignedLicenses).Count) { 'Leaver: unlicensed' }
        if ($why -and -not $wanted.Contains("$($d.Url)".TrimEnd('/'))) { $wanted["$($d.Url)".TrimEnd('/')] = @{ Upn = $owner; Why = $why; Site = $d } }
    }
}

# ------------------------------------------------------------ rows
$rows = foreach ($url in $wanted.Keys) {
    $w = $wanted[$url]
    $site = $w.Site
    $upn = $w.Upn
    $status = 'Copy'; $reason = $w.Why
    # Archive and lock state are only reliable on a per-site read.
    try { $full = Get-SPOSite -Identity $url -ErrorAction Stop } catch { $full = $site }
    $archive = "$($full.ArchiveStatus)"
    if ($archive -and $archive -ne 'NotArchived') { $status = 'Skip'; $reason = "ArchiveStatus $archive (Microsoft 365 Archive). Reactivate it before any copy tool can read it" }
    elseif ("$($full.LockState)" -eq 'NoAccess') { $status = 'Skip'; $reason = 'LockState NoAccess. Unlock it (Set-SPOSite -LockState Unlock) first' }

    $display = if ("$($full.Title)") { "$($full.Title)" } else { $upn }
    $destSite = $DestinationSite; $destLib = $DestinationLibrary; $folder = ''
    $mapped = if ($upn) { $map[$upn] } else { $null }
    if ($mapped) {
        $destSite = Get-PersonalUrl $dstTenant $mapped; $destLib = 'Documents'
        if ($status -eq 'Copy') { $reason = "$reason. Current staff: into their new OneDrive ($mapped)".TrimStart('. ') }
    }
    else {
        $local = if ($upn) { $upn.Split('@')[0] } else { ([uri]$url).Segments[-1] }
        $folder = Get-SafeFolder $(if ($FolderName -eq 'DisplayName') { $display } else { $local })
        if (-not $destSite -and $status -eq 'Copy') { $status = 'Skip'; $reason = 'No destination: not in the mapping file and no -DestinationSite' }
    }
    [pscustomobject]@{
        Source             = $url
        DestinationSite    = $destSite
        DestinationLibrary = $destLib
        DestinationFolder  = $folder
        SourceLibrary      = 'Documents'
        Status             = $status
        Reason             = $reason
        Owner              = $upn
        DisplayName        = $display
        StorageGB          = [math]::Round([double]$full.StorageUsageCurrent / 1024, 2)
        LastModified       = $full.LastContentModifiedDate
    }
}
$rows = @($rows)
foreach ($n in $notFound) {
    $rows += [pscustomobject]@{ Source = ''; DestinationSite = ''; DestinationLibrary = ''; DestinationFolder = ''; SourceLibrary = ''
        Status = 'Skip'; Reason = $n[1]; Owner = $n[0]; DisplayName = ''; StorageGB = 0; LastModified = '' }
}

# Two people with the same folder name would land in one folder.
$dupes = @($rows | Where-Object { $_.Status -eq 'Copy' -and $_.DestinationFolder } | Group-Object DestinationSite, DestinationLibrary, DestinationFolder | Where-Object Count -gt 1)
foreach ($g in $dupes) { foreach ($r in $g.Group) { $r.Status = 'Skip'; $r.Reason = "Folder name '$($r.DestinationFolder)' is used by $($g.Count) people. Edit DestinationFolder, then set Status to Copy" } }

$rows | Export-Csv $OutFile -NoTypeInformation -WhatIf:$false
$copy = @($rows | Where-Object Status -eq 'Copy')
Write-Host ''
$rows | Format-Table Owner, Status, StorageGB, DestinationFolder, Reason -AutoSize -Wrap | Out-String -Width 250 | Write-Host
Write-Host ("{0} to copy ({1} GB), {2} skipped. Written to {3}" -f $copy.Count, [math]::Round(($copy | Measure-Object StorageGB -Sum).Sum, 2), ($rows.Count - $copy.Count), $OutFile) -ForegroundColor Green

# ------------------------------------------------------------ folders
if ($CreateFolders) {
    $folders = @($copy | Where-Object { $_.DestinationFolder -and $_.DestinationSite -eq $DestinationSite } | ForEach-Object { $_.DestinationFolder } | Sort-Object -Unique)
    if (-not $folders) { Write-Host 'No folders to create.' -ForegroundColor Yellow; return }
    $graphTenant = if ($DestinationTenantId) { $DestinationTenantId } else { "$dstTenant.onmicrosoft.com" }
    Write-Host 'Sign in to the DESTINATION tenant for the folders.' -ForegroundColor Magenta
    Connect-GraphAs $graphTenant @('Sites.ReadWrite.All')
    $G = 'https://graph.microsoft.com/v1.0'
    $u = [uri]$DestinationSite.TrimEnd('/')
    $siteId = (Invoke-MgGraphRequest GET "$G/sites/$($u.Host):$($u.AbsolutePath)").id
    $libs = (Invoke-MgGraphRequest GET "$G/sites/$siteId/drives").value
    $drive = $libs | Where-Object { $_.name -eq $DestinationLibrary } | Select-Object -First 1
    if (-not $drive) { throw "No library '$DestinationLibrary' on $DestinationSite. Found: $(($libs | ForEach-Object { $_.name }) -join ', ')" }
    foreach ($f in $folders) {
        if (-not $PSCmdlet.ShouldProcess("$DestinationLibrary/$f", 'Create folder')) { continue }
        $body = @{ name = $f; folder = @{}; '@microsoft.graph.conflictBehavior' = 'fail' } | ConvertTo-Json
        try { Invoke-MgGraphRequest POST "$G/drives/$($drive.id)/root/children" -Body $body -ContentType 'application/json' | Out-Null; Write-Host "  Created $f" -ForegroundColor Green }
        catch {
            if ("$($_.Exception.Message)$($_.ErrorDetails.Message)" -match 'nameAlreadyExists|409|Conflict') { Write-Host "  Exists  $f" -ForegroundColor Gray }
            else { Write-Warning "Couldn't create $f : $($_.Exception.Message)" }
        }
    }
}
Write-Host 'Next: open the CSV, check every Skip, then run B-ShareGate\Copy-DrivesWithShareGate.ps1 -CsvPath (5.1) or C-No-budget\Copy-DriveViaGraph.ps1 -CsvPath.' -ForegroundColor Cyan
