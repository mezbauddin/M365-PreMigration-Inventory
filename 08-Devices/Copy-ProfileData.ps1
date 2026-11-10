<#
.SYNOPSIS
    On the laptop itself: copies a person's files and key settings from their old Windows profile into their new one
    after the laptop has moved tenants. Never changes or deletes the old profile.

.DESCRIPTION
    TL;DR: When a laptop leaves one tenant and joins another, the person signs in with a new account and Windows gives
    them a new, empty profile. Their old files are still on the disk, in the old profile folder. This script copies them
    across: Desktop, Documents, Downloads, Pictures, Videos, Music, Favorites, the old OneDrive folder(s), Outlook
    signatures, Sticky Notes, and Edge and Chrome bookmarks. The old profile is left exactly as it was, as a fallback.

    Two modes:
      -Prepare   Run BEFORE the laptop leaves the old tenant (as an admin). Saves the BitLocker recovery key, the
                 join status, printers and a profile list to C:\MigrationBackup on the laptop. Changes nothing else.
      (default)  Run AFTER the person has signed in once with the new account. Copies old to new.

    The new profile is the one belonging to -NewUser (default: the most recently created profile). The old profile is
    -OldProfile (a folder under C:\Users), or picked from a list.

    What it can't bring across: saved browser passwords (they're encrypted to the old account; export them first or
    re-enter them), and OneDrive files that were cloud-only (not downloaded) when the old account stopped syncing.
    Those are still in the old OneDrive in the cloud and come across with the OneDrive copy.

.PARAMETER Prepare
    Pre-move snapshot only. Saves BitLocker key, dsregcmd status, printers and profile list to -BackupFolder.

.PARAMETER OldProfile
    Old profile folder name or path, e.g. jsmith or C:\Users\jsmith. If omitted, the script lists profiles and asks.

.PARAMETER NewProfile
    New profile folder. Default: the newest profile folder created after the old one.

.PARAMETER BackupFolder
    Where -Prepare writes and where logs go. Default C:\MigrationBackup.

.PARAMETER WhatIf
    Show what would be copied without copying.

.EXAMPLE
    # Before leaving the old tenant, in an elevated Windows PowerShell on the laptop
    .\Copy-ProfileData.ps1 -Prepare

.EXAMPLE
    # After the person has signed in once with their new account, elevated again
    .\Copy-ProfileData.ps1 -OldProfile jsmith -WhatIf
    .\Copy-ProfileData.ps1 -OldProfile jsmith

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.0.1
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: Windows PowerShell 5.1 or PowerShell 7, run as administrator on the laptop. Uses robocopy (built in).
    Close Outlook, Edge and Chrome in the new profile before the copy.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$Prepare,
    [string]$OldProfile,
    [string]$NewProfile,
    [string]$BackupFolder = 'C:\MigrationBackup'
)

$ErrorActionPreference = 'Stop'
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) { throw 'Run this in an elevated PowerShell (Run as administrator).' }
New-Item -ItemType Directory -Path $BackupFolder -Force -WhatIf:$false | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmm'

function Get-Profiles {
    Get-CimInstance Win32_UserProfile | Where-Object { -not $_.Special -and $_.LocalPath -like 'C:\Users\*' } |
        ForEach-Object {
            $sid = $_.SID
            $name = try { ([Security.Principal.SecurityIdentifier]$sid).Translate([Security.Principal.NTAccount]).Value } catch { '(unknown account)' }
            $created = (Get-Item $_.LocalPath -ErrorAction SilentlyContinue).CreationTime
            [pscustomobject]@{ Folder = Split-Path $_.LocalPath -Leaf; Path = $_.LocalPath; Account = $name; Sid = $sid
                               Created = $created; LastUsed = $_.LastUseTime; Loaded = $_.Loaded }
        } | Sort-Object Created
}

# ---------- Prepare: before the laptop leaves the old tenant ----------
if ($Prepare) {
    $out = Join-Path $BackupFolder "Prepare-$env:COMPUTERNAME-$stamp.txt"
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("Computer: $env:COMPUTERNAME   Taken: $(Get-Date)")
    $lines.Add('===== BitLocker (C:) =====')
    try { $lines.AddRange([string[]](manage-bde -protectors -get C: 2>&1)) } catch { $lines.Add("manage-bde failed: $_") }
    $lines.Add('===== Join status (dsregcmd /status) =====')
    $lines.AddRange([string[]](dsregcmd /status 2>&1 | Select-String 'AzureAdJoined|DomainJoined|TenantName|TenantId|DeviceId|WorkplaceJoined'))
    $lines.Add('===== Printers =====')
    try { Get-Printer | ForEach-Object { $lines.Add("$($_.Name) | $($_.PortName) | $($_.DriverName)") } } catch { $lines.Add('Get-Printer not available') }
    $lines.Add('===== Profiles =====')
    Get-Profiles | ForEach-Object { $lines.Add("$($_.Folder) | $($_.Account) | created $($_.Created) | last used $($_.LastUsed)") }
    $lines | Set-Content -Path $out -Encoding UTF8 -WhatIf:$false
    Write-Host "Saved: $out" -ForegroundColor Green
    if (-not ($lines -match 'Numerical Password|Password:')) {
        Write-Warning 'No BitLocker recovery password found. Either BitLocker is off on C: (fine) or it has no recovery password. Check before leaving the tenant.'
    } else {
        Write-Warning "The file holds the BitLocker recovery key. Photograph it or copy it into your password manager, then keep $BackupFolder private."
    }
    Write-Host 'Next: create a local admin account if there isn''t one, disconnect from the old tenant, join the new one, have the person sign in once, then run this script again without -Prepare.' -ForegroundColor Cyan
    return
}

# ---------- Copy: after the person has signed in with the new account ----------
$profiles = @(Get-Profiles)
if (-not $OldProfile) {
    $profiles | Format-Table Folder, Account, Created, LastUsed -AutoSize | Out-String -Width 200 | Write-Host
    $OldProfile = Read-Host 'Type the OLD profile folder name (from the Folder column)'
}
$old = $profiles | Where-Object { $_.Folder -eq (Split-Path $OldProfile -Leaf) } | Select-Object -First 1
if (-not $old) { throw "No profile folder '$OldProfile' under C:\Users." }

if ($NewProfile) { $new = $profiles | Where-Object { $_.Folder -eq (Split-Path $NewProfile -Leaf) } | Select-Object -First 1 }
else { $new = $profiles | Where-Object { $_.Created -gt $old.Created -and $_.Path -ne $old.Path } | Sort-Object Created | Select-Object -Last 1 }
if (-not $new) { throw 'No new profile found. Have the person sign in once with their new account first, then run this again.' }
if ($new.Path -eq $old.Path) { throw 'Old and new profile are the same folder. Nothing to do.' }

Write-Host "Old profile: $($old.Path)  ($($old.Account))" -ForegroundColor Yellow
Write-Host "New profile: $($new.Path)  ($($new.Account))" -ForegroundColor Green
$running = Get-Process -Name outlook, msedge, chrome -ErrorAction SilentlyContinue
if ($running) { Write-Warning "Close these first for a clean copy: $(($running.Name | Sort-Object -Unique) -join ', ')" }

$log = Join-Path $BackupFolder "ProfileCopy-$env:COMPUTERNAME-$stamp.log"
$report = [System.Collections.Generic.List[object]]::new()
$oldLabel = "From old profile ($($old.Folder))"

function Copy-Tree($From, $To, $What) {
    if (-not (Test-Path $From)) { return }
    $all = @(Get-ChildItem $From -Recurse -File -Force -ErrorAction SilentlyContinue)
    # Cloud-only OneDrive placeholders (Offline / recall-on-access) show a full size but have no data on disk:
    # count them separately so the report doesn't promise files that can't be copied from this machine
    $cloudOnly = @($all | Where-Object { ([int]$_.Attributes -band 0x1000) -or ([int]$_.Attributes -band 0x400000) -or ([int]$_.Attributes -band 0x40000) })
    $files = @($all | Where-Object { $cloudOnly -notcontains $_ })
    $size = [math]::Round((($files | Measure-Object Length -Sum).Sum) / 1MB, 1)
    if ($PSCmdlet.ShouldProcess("$What ($($files.Count) files, $size MB)", "Copy to $To")) {
        # /XA:O skips cloud-only OneDrive placeholders (they can't download now and are in the cloud copy anyway)
        # No /COPY:S, so copied files inherit the new profile's permissions
        robocopy "$From" "$To" /E /COPY:DAT /DCOPY:T /XJ /XA:O /R:1 /W:1 /MT:8 /NP /NFL /NDL "/LOG+:$log" | Out-Null
        $code = $LASTEXITCODE
        $status = if ($code -lt 8) { 'Copied' } else { "Problem (robocopy $code, see log)" }
    } else { $status = 'WhatIf'; $code = $null }
    $report.Add([pscustomobject]@{ Item = $What; Files = $files.Count; MB = $size; CloudOnlyNotCopied = $cloudOnly.Count; Status = $status; To = $To })
    if ($cloudOnly.Count) { Write-Warning "$What : $($cloudOnly.Count) cloud-only file(s) were not on this laptop and were not copied. They come across with the OneDrive copy (ShareGate or Graph), not from here." }
}

foreach ($f in 'Desktop', 'Documents', 'Downloads', 'Pictures', 'Videos', 'Music', 'Favorites') {
    Copy-Tree (Join-Path $old.Path $f) (Join-Path $new.Path $f) $f
}
# Old OneDrive folder(s): kept separate so they don't mix with the new OneDrive while it syncs
Get-ChildItem $old.Path -Directory -Filter 'OneDrive*' -Force -ErrorAction SilentlyContinue | ForEach-Object {
    Copy-Tree $_.FullName (Join-Path $new.Path "$oldLabel\$($_.Name)") "Old $($_.Name) folder"
}
# Outlook signatures
Copy-Tree (Join-Path $old.Path 'AppData\Roaming\Microsoft\Signatures') (Join-Path $new.Path 'AppData\Roaming\Microsoft\Signatures') 'Outlook signatures'
# Outlook data files (.pst) wherever they are in the old profile
$psts = @(Get-ChildItem $old.Path -Recurse -Filter *.pst -File -Force -ErrorAction SilentlyContinue | Where-Object { $_.FullName -notlike '*\OneDrive*' })
foreach ($p in $psts) {
    $dest = Join-Path $new.Path "$oldLabel\Outlook data files"
    if ($PSCmdlet.ShouldProcess($p.FullName, "Copy to $dest")) {
        New-Item -ItemType Directory -Path $dest -Force | Out-Null; Copy-Item $p.FullName $dest -Force
        $report.Add([pscustomobject]@{ Item = "PST $($p.Name)"; Files = 1; MB = [math]::Round($p.Length / 1MB, 1); Status = 'Copied (open it in Outlook: File > Open > Outlook Data File)'; To = $dest })
    } else { $report.Add([pscustomobject]@{ Item = "PST $($p.Name)"; Files = 1; MB = [math]::Round($p.Length / 1MB, 1); Status = 'WhatIf'; To = $dest }) }
}
# Sticky Notes
$sn = 'AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState'
Copy-Tree (Join-Path $old.Path $sn) (Join-Path $new.Path "$oldLabel\Sticky Notes") 'Sticky Notes (database, kept aside)'
# Browser bookmarks: put straight in if the new browser profile has none yet, otherwise keep aside
foreach ($b in @(@{ N = 'Edge'; P = 'AppData\Local\Microsoft\Edge\User Data\Default\Bookmarks' }, @{ N = 'Chrome'; P = 'AppData\Local\Google\Chrome\User Data\Default\Bookmarks' })) {
    $src = Join-Path $old.Path $b.P
    if (-not (Test-Path $src)) { continue }
    $dst = Join-Path $new.Path $b.P
    $target = if (Test-Path $dst) { Join-Path $new.Path "$oldLabel\$($b.N) bookmarks (Bookmarks file)" } else { $dst }
    if ($PSCmdlet.ShouldProcess("$($b.N) bookmarks", "Copy to $target")) {
        New-Item -ItemType Directory -Path (Split-Path $target) -Force | Out-Null; Copy-Item $src $target -Force
        $st = if ($target -eq $dst) { 'Copied into the new browser profile' } else { 'Kept aside: the new browser already has bookmarks' }
    } else { $st = 'WhatIf' }
    $report.Add([pscustomobject]@{ Item = "$($b.N) bookmarks"; Files = 1; MB = 0; Status = $st; To = $target })
}

$csv = Join-Path $BackupFolder "ProfileCopy-$env:COMPUTERNAME-$stamp.csv"
$report | Export-Csv $csv -NoTypeInformation -WhatIf:$false
$report | Format-Table Item, Files, MB, Status -AutoSize | Out-String -Width 220 | Write-Host
Write-Host "Report: $csv   Log: $log" -ForegroundColor Cyan
Write-Host "The old profile ($($old.Path)) is untouched. Keep it for a few weeks before deleting it (Settings > System > About > Advanced system settings > User Profiles)." -ForegroundColor Cyan
Write-Host 'Not copied: saved browser passwords (re-enter or import from an export) and cloud-only OneDrive files (they arrive with the OneDrive copy).' -ForegroundColor Yellow
