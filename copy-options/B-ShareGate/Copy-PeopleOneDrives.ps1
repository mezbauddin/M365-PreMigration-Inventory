<#
.SYNOPSIS
    Copies people's OneDrives to their new OneDrives with ShareGate, end to end: finds every real OneDrive URL in
    both tenants, creates missing destination OneDrives, gives your admin access, then runs the copy.

.DESCRIPTION
    TL;DR: Give it the user names and your user mapping. It does the fiddly parts nobody should do by hand:

      1. Source: finds each person's OneDrive by owner (not by guessing the URL, which changes after a sign-in
         rename). No OneDrive (never used), archived or locked: reported and skipped. Adds you as site admin.
      2. Destination: finds the person's new OneDrive by owner. Missing: asks SharePoint to create it
         (Request-SPOPersonalSite) and waits. Adds you as site admin.
      3. Writes the copy list (CSV) and runs Copy-DrivesWithShareGate.ps1 with your .sgum mapping, optionally as an
         incremental, or as ShareGate's pre-check with -WhatIf.

    Why owner and not URL: OneDrive URLs are built from the sign-in name at the time the OneDrive was created, and
    SharePoint renames them later, in the background. After a cutover the guessed URL is often wrong in one tenant or
    the other, and ShareGate fails with CO-001-032 ("The connection to the SharePoint site failed").

.PARAMETER MappingFile
    UserMapping.csv from 02b (columns Source, Destination). Rebuild it after cutover: sign-ins change.

.PARAMETER Users
    User names (aperez) or source sign-ins (aperez@fabrikam.onmicrosoft.com). Matched against the Source column.

.PARAMETER SourceAdminUrl / DestinationAdminUrl
    SharePoint admin URLs, e.g. https://fabrikam-admin.sharepoint.com

.PARAMETER SourceAdmin / DestinationAdmin
    Your admin sign-in in each tenant. Added as site admin on each OneDrive so ShareGate can read and write.

.PARAMETER SgumFile
    ShareGate user mapping (.sgum) from 02b, so authors and sharing land on the right people.

.PARAMETER Incremental
    Copy only new or changed items (run again later, or after a first full copy).

.PARAMETER PrepareOnly
    Do steps 1 and 2 and write the list, but don't copy.

.PARAMETER ProvisionWaitMinutes
    How long to wait for newly requested destination OneDrives. Default 20. Not ready by then: run again later.

.PARAMETER WhatIf
    Steps 1 and 2 still run (they only add admin rights and create OneDrives); the copy is ShareGate's pre-check.

.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File .\Copy-PeopleOneDrives.ps1 -MappingFile .\Mapping-PostCutover\UserMapping.csv `
        -Users aperez, jmcevoy, ktaylor -SourceAdminUrl https://fabrikam-admin.sharepoint.com -SourceAdmin admin@fabrikam.onmicrosoft.com `
        -DestinationAdminUrl https://contoso-admin.sharepoint.com -DestinationAdmin admin@contoso.com -SgumFile .\Mapping-PostCutover\UserMapping.sgum

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.0.0
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: Windows PowerShell 5.1 (ShareGate's module and the SharePoint Online module both load there),
              ShareGate Migrate Pro or Enterprise, Microsoft.Online.SharePoint.PowerShell, SharePoint admin in both
              tenants. Not sponsored by ShareGate.
    Changelog:
      1.0.0  2026-11-10  First release, from a live cutover where guessed OneDrive URLs failed in both tenants.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$MappingFile,
    [Parameter(Mandatory)][string[]]$Users,
    [Parameter(Mandatory)][string]$SourceAdminUrl,
    [Parameter(Mandatory)][string]$SourceAdmin,
    [Parameter(Mandatory)][string]$DestinationAdminUrl,
    [Parameter(Mandatory)][string]$DestinationAdmin,
    [string]$SgumFile,
    [switch]$Incremental,
    [switch]$PrepareOnly,
    [int]$ProvisionWaitMinutes = 20,
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSEdition -eq 'Core') {
    throw 'Run this in Windows PowerShell 5.1: powershell.exe -ExecutionPolicy Bypass -File .\Copy-PeopleOneDrives.ps1 <same parameters>'
}
if (-not (Get-Module -ListAvailable Microsoft.Online.SharePoint.PowerShell)) {
    Write-Host 'Installing the SharePoint Online module (current user)...' -ForegroundColor Cyan
    Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -Force -AllowClobber
}
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -WarningAction SilentlyContinue

$stamp = Get-Date -Format 'yyyyMMdd-HHmm'
$map = @(Import-Csv $MappingFile)
foreach ($c in 'Source', 'Destination') { if (-not ($map | Get-Member -Name $c)) { throw "$MappingFile needs a '$c' column." } }

# Who
$people = foreach ($u in $Users) {
    $u = $u.Trim().ToLower()
    $row = @($map | Where-Object { $_.Source.ToLower() -eq $u -or $_.Source.ToLower() -like "$u@*" })
    if ($row.Count -ne 1) { Write-Warning "$u : $($row.Count) matches in $MappingFile, skipped."; continue }
    if (-not $row[0].Destination) { Write-Warning "$u : no destination in the mapping, skipped."; continue }
    [pscustomobject]@{ User = $u.Split('@')[0]; SourceUpn = $row[0].Source; DestinationUpn = $row[0].Destination
                       Source = ''; DestinationSite = ''; SourceMB = $null; Status = 'Copy'; Reason = '' }
}
if (-not $people) { throw 'Nobody to copy.' }

function Get-PersonalSites { @(Get-SPOSite -IncludePersonalSite $true -Limit All -Filter "Url -like '-my.sharepoint.com/personal/'") }
function Find-Drive($sites, [string]$upn) {
    $hit = @($sites | Where-Object { "$($_.Owner)".ToLower() -eq $upn.ToLower() })
    if (-not $hit) { $seg = ($upn.ToLower() -replace '[@.]', '_'); $hit = @($sites | Where-Object { $_.Url.TrimEnd('/').ToLower().EndsWith("/$seg") }) }
    if ($hit.Count -eq 1) { $hit[0] } else { $null }
}
function Grant-Admin([string]$url, [string]$admin) {
    Set-SPOUser -Site $url -LoginName $admin -IsSiteCollectionAdmin $true | Out-Null
}

# 1. Source
Write-Host "`n== Source: sign in as $SourceAdmin" -ForegroundColor Magenta
Connect-SPOService -Url $SourceAdminUrl
$src = Get-PersonalSites
Write-Host "$($src.Count) OneDrive(s) in the source." -ForegroundColor Cyan
foreach ($p in $people) {
    $d = Find-Drive $src $p.SourceUpn
    if (-not $d) { $p.Status = 'Skip'; $p.Reason = 'No OneDrive in the source (never used): nothing to copy'; continue }
    $p.Source = $d.Url; $p.SourceMB = $d.StorageUsageCurrent
    $arch = $d.PSObject.Properties['ArchiveStatus']
    if ($arch -and $d.ArchiveStatus -and $d.ArchiveStatus -ne 'NotArchived') { $p.Status = 'Skip'; $p.Reason = "Archived ($($d.ArchiveStatus)): reactivate first"; continue }
    if ($d.LockState -eq 'NoAccess') { $p.Status = 'Skip'; $p.Reason = 'Locked (NoAccess): unlock first'; continue }
    try { Grant-Admin $d.Url $SourceAdmin } catch { $p.Status = 'Skip'; $p.Reason = "Couldn't add you as admin: $($_.Exception.Message)" }
}
Disconnect-SPOService -ErrorAction SilentlyContinue

# 2. Destination
$todo = @($people | Where-Object Status -eq 'Copy')
if ($todo) {
    Write-Host "`n== Destination: sign in as $DestinationAdmin" -ForegroundColor Magenta
    Connect-SPOService -Url $DestinationAdminUrl
    $dst = Get-PersonalSites
    $missing = @()
    foreach ($p in $todo) {
        $d = Find-Drive $dst $p.DestinationUpn
        if ($d) { $p.DestinationSite = $d.Url } else { $missing += $p }
    }
    if ($missing) {
        Write-Host "Creating $($missing.Count) destination OneDrive(s): $(($missing.DestinationUpn) -join ', ')" -ForegroundColor Cyan
        Request-SPOPersonalSite -UserEmails @($missing.DestinationUpn) -NoWait
        $until = (Get-Date).AddMinutes($ProvisionWaitMinutes)
        while (@($missing | Where-Object { -not $_.DestinationSite }).Count -and (Get-Date) -lt $until) {
            Start-Sleep -Seconds 60
            $dst = Get-PersonalSites
            foreach ($p in $missing) { if (-not $p.DestinationSite) { $d = Find-Drive $dst $p.DestinationUpn; if ($d) { $p.DestinationSite = $d.Url; Write-Host "   Ready: $($p.DestinationUpn)" -ForegroundColor Green } } }
        }
        foreach ($p in $missing) { if (-not $p.DestinationSite) { $p.Status = 'Skip'; $p.Reason = "Destination OneDrive still being created (asked $ProvisionWaitMinutes min ago). Run again later." } }
    }
    foreach ($p in @($todo | Where-Object { $_.Status -eq 'Copy' })) {
        try { Grant-Admin $p.DestinationSite $DestinationAdmin } catch { $p.Status = 'Skip'; $p.Reason = "Couldn't add you as admin: $($_.Exception.Message)" }
    }
    Disconnect-SPOService -ErrorAction SilentlyContinue
}

# 3. List and copy
$list = ".\PeopleDrives-$stamp.csv"
$people | Select-Object User, SourceUpn, DestinationUpn, Source, DestinationSite, @{ n = 'SourceGB'; e = { if ($_.SourceMB) { [math]::Round($_.SourceMB / 1024, 2) } } }, Status, Reason |
    Export-Csv $list -NoTypeInformation
Write-Host ''
$people | Format-Table User, Status, @{ n = 'GB'; e = { if ($_.SourceMB) { [math]::Round($_.SourceMB / 1024, 2) } } }, Reason -AutoSize -Wrap | Out-String -Width 250 | Write-Host
Write-Host "Copy list: $list (holds names and URLs: keep it out of source control)." -ForegroundColor Cyan

if ($PrepareOnly) { return }
if (-not @($people | Where-Object Status -eq 'Copy')) { Write-Warning 'Nothing to copy.'; return }
$engine = Join-Path $PSScriptRoot 'Copy-DrivesWithShareGate.ps1'
if (-not (Test-Path $engine)) { throw "Copy-DrivesWithShareGate.ps1 not found next to this script ($PSScriptRoot)." }
$a = @{ CsvPath = $list; SourceAdminUrl = $SourceAdminUrl; DestinationAdminUrl = $DestinationAdminUrl }
if ($SgumFile) { $a.MappingFile = $SgumFile }
if ($Incremental) { $a.Incremental = $true }
if ($WhatIf) { $a.WhatIf = $true }
& $engine @a
