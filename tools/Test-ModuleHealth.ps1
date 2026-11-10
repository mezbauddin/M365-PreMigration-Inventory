<#
.SYNOPSIS
    Checks the PowerShell modules this toolkit needs for the version clashes that break a migration session, and
    optionally fixes the Microsoft Graph ones.

.DESCRIPTION
    TL;DR: Run this before you start, and again whenever a script fails with "Could not load file or assembly" or
    "Assembly with same name is already loaded". It lists every installed version of Microsoft.Graph.Authentication
    and the other Microsoft.Graph.* modules, ExchangeOnlineManagement, Microsoft.Online.SharePoint.PowerShell,
    PnP.PowerShell and MicrosoftTeams, then flags:
      - Graph submodules whose newest version differs from the newest Microsoft.Graph.Authentication (the classic
        clash: Authentication 2.41.2 loaded, then an older submodule asks for Authentication 2.26.1)
      - more than one ExchangeOnlineManagement version installed
      - Exchange Online or Graph assemblies already loaded in this window (open a fresh window before a run)
      - the PowerShell edition, and whether ShareGate's module is available (it only loads in Windows PowerShell 5.1)
    Read-only by default. -Fix installs each mismatched Microsoft.Graph.* module at the Authentication version, for
    the current user only.

.PARAMETER Fix
    Install-Module every mismatched Microsoft.Graph.* submodule at the newest Microsoft.Graph.Authentication version
    (-Scope CurrentUser -Force -AllowClobber). Asks before each one; supports -WhatIf. Close every PowerShell window
    afterwards and open a fresh one.

.PARAMETER OutFile
    Optional CSV path for the result table.

.EXAMPLE
    .\Test-ModuleHealth.ps1

.EXAMPLE
    .\Test-ModuleHealth.ps1 -Fix -WhatIf

.EXAMPLE
    .\Test-ModuleHealth.ps1 -Fix -Confirm:$false

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.0.0
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: PowerShell 5.1 or 7. PowerShellGet for -Fix and for Get-InstalledModule. No tenant sign-in.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$Fix,
    [string]$OutFile
)

$ErrorActionPreference = 'Stop'
$results = [System.Collections.Generic.List[object]]::new()
function Add-Check($Area, $Item, $Result, $Detail) {
    $results.Add([pscustomobject]@{ Area = $Area; Item = $Item; Result = $Result; Detail = $Detail })
}
function Get-Versions($Name) {
    # Both sources: Get-Module sees every copy on PSModulePath; Get-InstalledModule only what PowerShellGet installed.
    $v = @(Get-Module -ListAvailable -Name $Name -ErrorAction SilentlyContinue | ForEach-Object { $_.Version })
    if (Get-Command Get-InstalledModule -ErrorAction SilentlyContinue) {
        $v += @(Get-InstalledModule -Name $Name -AllVersions -ErrorAction SilentlyContinue | ForEach-Object { [version]("$($_.Version)" -replace '-.*$', '') })
    }
    @($v | Where-Object { $_ } | Sort-Object -Unique -Descending)
}

# ------------------------------------------------------------ PowerShell edition and ShareGate
$edition = "$($PSVersionTable.PSEdition) $($PSVersionTable.PSVersion)"
Add-Check 'PowerShell' 'Edition' 'PASS' "$edition. The toolkit runs in 7; ShareGate's module needs Windows PowerShell 5.1"
$sg = @(Get-Module -ListAvailable -Name Sharegate -ErrorAction SilentlyContinue)
if ($sg) {
    if ($PSVersionTable.PSEdition -eq 'Core') { Add-Check 'ShareGate' 'Sharegate module' 'WARN' "Found ($($sg[0].Version)), but this is PowerShell 7. Run ShareGate scripts with powershell.exe (5.1)" }
    else { Add-Check 'ShareGate' 'Sharegate module' 'PASS' "Found ($($sg[0].Version)) in Windows PowerShell 5.1" }
}
else { Add-Check 'ShareGate' 'Sharegate module' 'PASS' 'Not visible here. Only needed for copy-options/B-ShareGate scripts; it installs with ShareGate Migrate and loads in 5.1 only' }

# ------------------------------------------------------------ Microsoft Graph
$authVersions = Get-Versions 'Microsoft.Graph.Authentication'
$authTop = $authVersions | Select-Object -First 1
$graphNames = @(Get-Module -ListAvailable -Name 'Microsoft.Graph.*' -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })
if (Get-Command Get-InstalledModule -ErrorAction SilentlyContinue) {
    $graphNames += @(Get-InstalledModule -Name 'Microsoft.Graph.*' -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })
}
$graphNames = @($graphNames | Where-Object { $_ -ne 'Microsoft.Graph.Authentication' -and $_ -ne 'Microsoft.Graph' } | Sort-Object -Unique)

$mismatched = [System.Collections.Generic.List[string]]::new()
if (-not $authTop) {
    Add-Check 'Graph' 'Microsoft.Graph.Authentication' 'WARN' 'Not installed. Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
}
else {
    $r = if ($authVersions.Count -gt 1) { 'WARN' } else { 'PASS' }
    $d = "Installed: $($authVersions -join ', ')"
    if ($authVersions.Count -gt 1) { $d += '. More than one version: PowerShell loads the newest unless told otherwise, and older submodules then fail' }
    Add-Check 'Graph' 'Microsoft.Graph.Authentication' $r $d
    foreach ($n in $graphNames) {
        $v = Get-Versions $n
        $top = $v | Select-Object -First 1
        if ($top -ne $authTop) {
            $mismatched.Add($n)
            Add-Check 'Graph' $n 'WARN' "Newest $top, Authentication is $authTop. This is the 'Could not load file or assembly Microsoft.Graph.Authentication' clash. Fix: install $n $authTop (-Fix)"
        }
        else { Add-Check 'Graph' $n 'PASS' "Newest $top matches Authentication$(if ($v.Count -gt 1) { ". Older copies too: $(($v | Select-Object -Skip 1) -join ', ')" })" }
    }
}

# ------------------------------------------------------------ the other modules
$others = [ordered]@{
    'ExchangeOnlineManagement'               = 'Parts 01 to 09'
    'Microsoft.Online.SharePoint.PowerShell' = 'Parts 01, 04, 06, 09 and New-DriveCopyList'
    'PnP.PowerShell'                         = 'Optional: New-DriveCopyList -CreateFolders'
    'MicrosoftTeams'                         = 'Optional: checking Teams by hand'
}
foreach ($n in $others.Keys) {
    $v = Get-Versions $n
    if (-not $v) { Add-Check 'Modules' $n 'PASS' "Not installed. Used by: $($others[$n])"; continue }
    if ($n -eq 'ExchangeOnlineManagement' -and $v.Count -gt 1) {
        Add-Check 'Modules' $n 'WARN' "Several versions installed: $($v -join ', '). Different windows can load different ones. Keep one: Uninstall-Module ExchangeOnlineManagement -RequiredVersion <old>"
    }
    else { Add-Check 'Modules' $n 'PASS' "Installed: $($v -join ', ')" }
}

# ------------------------------------------------------------ what this window has already loaded
$loaded = @([AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object { $_.GetName() } |
        Where-Object { $_.Name -like 'Microsoft.Graph*' -or $_.Name -like 'Microsoft.Exchange.Management*' -or $_.Name -eq 'Microsoft.Identity.Client' })
if ($loaded) {
    $list = ($loaded | Sort-Object Name -Unique | ForEach-Object { "$($_.Name) $($_.Version)" }) -join '; '
    Add-Check 'Session' 'Loaded assemblies' 'WARN' "Already loaded in this window: $list. Exchange Online and Graph share Microsoft.Identity.Client, and whichever loads first wins. Run the next script in a fresh window"
}
else { Add-Check 'Session' 'Loaded assemblies' 'PASS' 'No Exchange Online or Graph assemblies loaded yet. Good: this window is clean' }

# ------------------------------------------------------------ fix
if ($Fix) {
    if (-not $mismatched.Count) { Write-Host 'Nothing to fix: every Microsoft.Graph submodule matches Authentication.' -ForegroundColor Green }
    foreach ($n in $mismatched) {
        if ($PSCmdlet.ShouldProcess("$n $authTop", 'Install-Module -Scope CurrentUser -Force -AllowClobber')) {
            try {
                Install-Module -Name $n -RequiredVersion $authTop -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
                Add-Check 'Fix' $n 'PASS' "Installed $authTop"
            }
            catch { Add-Check 'Fix' $n 'WARN' "Install failed: $($_.Exception.Message)" }
        }
    }
    if ($mismatched.Count) { Write-Host 'Close every PowerShell window and open a fresh one before you run anything else.' -ForegroundColor Yellow }
}

# ------------------------------------------------------------ report
$results | Format-Table Result, Area, Item, Detail -AutoSize -Wrap | Out-String -Width 250 | Write-Host
if ($OutFile) { $results | Export-Csv $OutFile -NoTypeInformation -WhatIf:$false; Write-Host "Saved to $OutFile" -ForegroundColor Cyan }
$warn = @($results | Where-Object Result -eq 'WARN').Count
if ($warn) { Write-Host "$warn warning(s). Fix them before a migration run, or expect assembly errors part way through." -ForegroundColor Yellow }
else { Write-Host 'All PASS.' -ForegroundColor Green }
