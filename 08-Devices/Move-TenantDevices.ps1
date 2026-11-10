<#
.SYNOPSIS
    One interactive script for moving Windows devices, their Intune configuration and Autopilot registrations from
    one Microsoft 365 tenant to another, keeping every person's Windows profile and data in place.

.DESCRIPTION
    TL;DR: Run it with no parameters and pick from the menu. Some options run on your admin PC (build the enrolment
    package, copy Intune configuration, move Autopilot hashes). The rest run on each laptop as administrator (safety
    net, leave the old tenant, join the new one silently, hand the person's existing profile to their new account).
    Every tenant-changing step has -WhatIf. Nothing is ever wiped, retired or reset.

    Where to run what: options 1 to 3 (and 3 > move) on any admin PC, your own laptop included, because they only
    sign in to tenants. Options 4 to 6 only on the device being moved. Never run option 4 on your own PC: it takes
    the machine it runs on out of its tenant. Option 4 guards against that: pass -SourceTenantId and it stops if
    this machine is joined anywhere else; leave it out and you must type the joined tenant's name exactly. It also
    stops before leaving unless the local admin safety-net account exists, is enabled and is an administrator.

    Menu
      1  Build the enrolment package      admin PC   Signs in to the destination, gets a bulk enrolment token, builds
                                                      a .ppkg that joins laptops silently. One package per tenant,
                                                      valid up to 180 days. No prompts: tenant and account come from
                                                      the sign-in. -AllowWizard falls back to the designer's wizard.
      2  Copy Intune configuration        admin PC   Exports policies, compliance, scripts, remediations, Autopilot
                                                      profiles and app metadata from the source, imports into the
                                                      destination, maps group assignments by name. -WhatIf first.
      3  Autopilot                        both       On a laptop: capture its hardware hash. On the admin PC: register
                                                      captured hashes in the destination and remove them from the source.
      4  Move this laptop                 laptop     Stage 1: pre-flight, BitLocker key, local admin, Autopilot hash,
                                                      profile list, leave the old tenant, restart. Stage 2 (after the
                                                      restart): install the .ppkg, join, restart. Stage 3: wait for the
                                                      person's first sign-in, then option 5. The script remembers the
                                                      stage in C:\MigrationBackup\state.json, so just run it again.
      5  Hand the old profile over        laptop     After the person has signed in once with the new account: points
                                                      their new account at the old profile folder, fixes ownership and
                                                      permissions, parks the empty new profile. Verified, with rollback.
      6  Copy profile data instead        laptop     Fallback: copies files, old OneDrive, signatures, PSTs, bookmarks
                                                      from the old profile into the new one (Copy-ProfileData.ps1).
      7  Clean up this admin PC           admin PC   Removes what option 1 installed or left behind: private designer
                                                      copies, temp files, the sign-in window setting, build logs, and
                                                      (if you confirm) Windows Configuration Designer and the ADK
                                                      designer. Keeps your .ppkg files. -WhatIf shows the list only.

.PARAMETER Action
    Skip the menu: BuildPackage, CopyIntune, AutopilotCapture, AutopilotMove, MoveDevice, HandoverProfile, CopyProfile,
    CleanupAdminPC.

.PARAMETER SourceTenantId
    Source tenant ID (GUID). CopyIntune and AutopilotMove sign in to it. MoveDevice uses it as a safety check: if
    this machine's dsregcmd TenantId is different, it stops before changing anything. Strongly recommended for
    MoveDevice.

.PARAMETER PackageName
    Option 1 only: package file name. Default: <tenant domain>-join-<date>.

.PARAMETER TokenDays
    Option 1 only: how long the package can join devices (1 to 180 days). Default 180.

.PARAMETER RenamePattern
    Option 1 only: rename laptops on join, for example TG-%SERIAL% (15 characters at most once expanded).
    Default: keep their names.

.PARAMETER AllowWizard
    Option 1 only: if the silent build fails, open the Windows Configuration Designer wizard instead of stopping.
    Off by default, so option 1 is no-touch.

.PARAMETER UseDeviceCode
    Option 1 only: sign in with a device code instead of the interactive sign-in. Off by default, because many
    tenants block device code flow with Conditional Access.

.PARAMETER WorkFolder
    Where state, keys and reports go. Default C:\MigrationBackup (laptops) or .\DeviceMigration (admin PC).

.EXAMPLE
    .\Move-TenantDevices.ps1
    # menu

.EXAMPLE
    .\Move-TenantDevices.ps1 -Action CopyIntune -SourceTenantId <guid> -DestinationTenantId <guid> -WhatIf

.EXAMPLE
    .\Move-TenantDevices.ps1 -Action MoveDevice -SourceTenantId <source-guid> -PackagePath C:\MigrationBackup\contoso-join.ppkg
    # on the laptop being moved, never on your own PC

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.2.5
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: Windows PowerShell 5.1 or 7, elevated. Admin-PC options need Microsoft.Graph.Authentication.
              Laptop options need nothing extra (Provisioning module is built in).
    Status  : Option 1's scripted token request uses the same sign-in flow Windows Configuration Designer uses, which
              Microsoft doesn't document. If it stops working, the script opens WCD for you instead. Options 2, 3 and 5
              are tested against stubbed tenants and devices; run them with -WhatIf first on anything that matters.
    Changelog:
      1.2.5  2026-11-10  Clearing old sign-ins no longer stops on a missing Office registry key (Windows PowerShell stderr).
      1.2.4  2026-11-10  Option 5 says how many files it is re-permissioning, that it takes a while, and when pass 1 ends.
      1.2.3  2026-11-10  Option 5: icacls 'Access is denied' on legacy junctions no longer aborts the handover under
                         Windows PowerShell; ownership verified instead. If a profile is still open, offers a restart.
                         New -Action Update (menu 8) pulls the latest version.
      1.2.2  2026-11-10  Option 5 clears the old tenant's cached sign-ins in the handed-over profile (sign-in broker, Office
                         identities, Teams cache), renamed or exported first so it can be undone.
      1.2.1  2026-11-10  Stage 2 finds the .ppkg in the work folder if it was copied in after stage 1. Generated local
                         admin password is short and easy to type on a call (Move1234!): the account is temporary.
      1.2.0  2026-11-10  Pick by number instead of typing: profiles, package, local admin password (can generate one),
                         leave / handover / restart confirmations. Entra profiles found by SID (S-1-12-1-*), so option 5
                         works when Windows shows a tenant prefix instead of AzureAD\. Package auto-found in the work folder.
      1.1.1  2026-11-10  Relaunches itself in 64-bit PowerShell when started from 32-bit (x86) PowerShell, where dsregcmd
                         is not found.
      1.1.0  2026-11-10  Option 7: clean up the admin PC (designer copies, temp files, sign-in setting, build logs, and on
                         confirmation Windows Configuration Designer and the ADK). Keeps .ppkg files.
      1.0.10 2026-11-10  Option 1 build: absolute work folder paths (ICD runs from its own folder); ADK designer run from a
                         private copy with the missing Microsoft.Diagnostics.Tracing.EventSource.dll added.
      1.0.9  2026-11-10  Option 1 sign-in no longer depends on MSAL files from other modules: plain OAuth code flow in a
                         sign-in window, then the system browser with PKCE.
      1.0.8  2026-11-10  Option 1 asks nothing: tenant, account and package name come from the sign-in (-PackageName,
                         -TokenDays, -RenamePattern to override). Sign-in window first (Windows broker removed: AADSTS50011).
      1.0.7  2026-11-10  Option 1 sign-in runs in a clean child PowerShell, so it works when Microsoft Graph is already
                         loaded in the session (its sign-in library is otherwise out of reach).
      1.0.6  2026-11-10  Option 1 is no-touch: builds with the ADK designer or a copy of the Store designer run from the
                         command line, installs the ADK designer silently if needed, reports exit codes and never opens
                         the wizard unless -AllowWizard.
      1.0.5  2026-11-10  Option 1 signs in interactively (Windows sign-in broker, then system browser, then sign-in
                         window) so MFA and Conditional Access apply. Device code only with -UseDeviceCode.
      1.0.4  2026-11-10  Option 1 builds the package with no clicks: real bulk token via bulkaadjtoken (the token
                         1.0.0-1.0.3 requested was an ordinary refresh token and would not have joined devices),
                         optional rename pattern, ICD run inside the Store app's package context, or the ADK
                         Configuration Designer installed silently with winget. Guided WCD stays as the last resort.
      1.0.3  2026-11-10  Guided build searches for the new .ppkg instead of assuming one folder; device name left blank.
      1.0.2  2026-11-10  Option 1 finds Windows Configuration Designer from the Microsoft Store as well as the ADK,
                         installs it with winget when missing, checks for it before asking you to sign in, and falls
                         back to a guided build when the Store version won't build from the command line.
      1.0.1  2026-11-10  MoveDevice safety guard: stops if -SourceTenantId is given and this machine is joined to a
                         different tenant; without it, shows the joined tenant name and ID and requires the tenant
                         name typed exactly (not just LEAVE). Stops before leaving unless the local admin exists, is
                         enabled and is in Administrators (checked by well-known SID, so it works on any language).
                         DeviceMove-*.csv gains a Status column for tools/Get-MigrationReport.ps1.
      1.0.0  2026-11-10  First release.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateSet('Menu', 'BuildPackage', 'CopyIntune', 'AutopilotCapture', 'AutopilotMove', 'MoveDevice', 'HandoverProfile', 'CopyProfile', 'CleanupAdminPC', 'Update')]
    [string]$Action = 'Menu',
    [string]$WorkFolder,
    [string]$PackagePath,
    [string]$OldProfile,
    [string]$NewProfile,
    [string]$SourceTenantId,
    [string]$DestinationTenantId,
    [string]$SourceAdmin,
    [string]$DestinationAdmin,
    [string]$ExportFolder,
    [string]$GroupMappingFile,
    [string]$LocalAdminName = 'migadmin',
    [string]$PackageName,
    [ValidateRange(1, 180)][int]$TokenDays = 180,
    [string]$RenamePattern,
    [switch]$UseDeviceCode,
    [switch]$AllowWizard,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
# 32-bit PowerShell on 64-bit Windows can't see dsregcmd, manage-bde and parts of the registry: relaunch as 64-bit
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    $ps64 = Join-Path $env:SystemRoot 'Sysnative\WindowsPowerShell\v1.0\powershell.exe'
    Write-Host 'This is 32-bit PowerShell (x86). Restarting in 64-bit PowerShell...' -ForegroundColor Yellow
    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
    foreach ($k in $PSBoundParameters.Keys) { $v = $PSBoundParameters[$k]; if ($v -is [switch]) { if ($v) { $argList += "-$k" } } else { $argList += "-$k"; $argList += "$v" } }
    & $ps64 @argList
    return
}
$script:Repo = 'https://raw.githubusercontent.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit/main/08-Devices'
if (-not $WorkFolder) { $WorkFolder = if (Test-Path 'C:\') { 'C:\MigrationBackup' } else { Join-Path (Get-Location) 'DeviceMigration' } }
New-Item -ItemType Directory -Path $WorkFolder -Force -WhatIf:$false | Out-Null
$WorkFolder = (Resolve-Path -LiteralPath $WorkFolder).Path   # absolute: tools run from their own folder
$script:StateFile = Join-Path $WorkFolder 'state.json'
$script:Stamp = Get-Date -Format 'yyyyMMdd-HHmm'

# ---------------------------------------------------------------- helpers
function Write-Step($t) { Write-Host "`n== $t" -ForegroundColor Cyan }
function Write-Ok($t) { Write-Host "   OK   $t" -ForegroundColor Green }
function Write-Warn2($t) { Write-Host "   WARN $t" -ForegroundColor Yellow }
function Write-Bad($t) { Write-Host "   FAIL $t" -ForegroundColor Red }
function Test-Elevated { ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) }
function Get-State { if (Test-Path $script:StateFile) { Get-Content $script:StateFile -Raw | ConvertFrom-Json } else { [pscustomobject]@{ Stage = 'NotStarted' } } }
function Set-State($obj) { $obj | ConvertTo-Json -Depth 5 | Set-Content $script:StateFile -Encoding UTF8 -WhatIf:$false }
function Add-State($name, $value) { $s = Get-State; $s | Add-Member -NotePropertyName $name -NotePropertyValue $value -Force; Set-State $s }
function Get-DsregStatus {
    $o = dsregcmd /status 2>&1 | Out-String
    $h = @{}
    foreach ($k in 'AzureAdJoined', 'DomainJoined', 'TenantName', 'TenantId', 'DeviceId', 'WorkplaceJoined', 'MdmUrl') {
        if ($o -match "$k\s*:\s*(.+)") { $h[$k] = $Matches[1].Trim() } else { $h[$k] = '' }
    }
    [pscustomobject]$h
}
function Test-LocalAdminReady([string]$Name) {
    # Returns $null when the safety-net account exists, is enabled and is a local administrator; otherwise the reason.
    $u = Get-LocalUser -Name $Name -ErrorAction SilentlyContinue
    if (-not $u) { return "Local admin '$Name' does not exist." }
    if (-not $u.Enabled) { return "Local admin '$Name' exists but is disabled." }
    $isAdmin = $false
    try {
        $isAdmin = [bool](Get-LocalGroupMember -SID 'S-1-5-32-544' -ErrorAction Stop | Where-Object { $_.SID -eq $u.SID })
    }
    catch {
        # Get-LocalGroupMember fails on groups holding orphaned SIDs; fall back to net localgroup
        $grp = ((New-Object System.Security.Principal.SecurityIdentifier 'S-1-5-32-544').Translate([System.Security.Principal.NTAccount]).Value -split '\\')[-1]
        $isAdmin = [bool]((net localgroup "$grp" 2>$null) | Where-Object { ($_ -split '\\')[-1].Trim() -eq $Name })
    }
    if (-not $isAdmin) { return "Local admin '$Name' is not a member of the local Administrators group." }
    $null
}
function Get-Profiles {
    Get-CimInstance Win32_UserProfile | Where-Object { -not $_.Special -and $_.LocalPath -like "$env:SystemDrive\Users\*" } | ForEach-Object {
        $sid = $_.SID
        $name = try { ([Security.Principal.SecurityIdentifier]$sid).Translate([Security.Principal.NTAccount]).Value } catch { '(unknown account)' }
        # Entra accounts have S-1-12-1-* SIDs. Their display name varies (AzureAD\name, or a tenant prefix like CONTOSO\name)
        [pscustomobject]@{ Folder = Split-Path $_.LocalPath -Leaf; Path = $_.LocalPath; Account = $name; Sid = $sid; IsEntra = ($sid -like 'S-1-12-1-*')
                           Created = (Get-Item $_.LocalPath -ErrorAction SilentlyContinue).CreationTime; LastUsed = $_.LastUseTime; Loaded = $_.Loaded }
    } | Sort-Object Created
}
function Connect-Tenant([string]$TenantId, [string]$Upn, [string[]]$Scopes, [string]$Label) {
    if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) { throw 'Install-Module Microsoft.Graph.Authentication -Scope CurrentUser first.' }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    $ctx = Get-MgContext
    if ($ctx -and (($ctx.TenantId -ne $TenantId) -or ($Upn -and $ctx.Account -ne $Upn))) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null; $ctx = $null }
    if (-not $ctx -or (@($Scopes | Where-Object { $ctx.Scopes -notcontains $_ }).Count -gt 0)) {
        Write-Host "Sign in to the $Label tenant$(if ($Upn) { " as $Upn" })." -ForegroundColor Magenta
        Connect-MgGraph -TenantId $TenantId -Scopes $Scopes -NoWelcome
        $ctx = Get-MgContext
    }
    if ($ctx.TenantId -ne $TenantId) { throw "Signed in to $($ctx.TenantId), not $TenantId. Disconnect-MgGraph and try again." }
    Write-Host "Signed in as $($ctx.Account) to tenant $($ctx.TenantId) ($Label)" -ForegroundColor Green
}
function Get-GraphAll($uri) {
    $out = [System.Collections.Generic.List[object]]::new()
    while ($uri) {
        $p = Invoke-MgGraphRequest GET $uri
        $isList = if ($p -is [System.Collections.IDictionary]) { $p.Contains('value') } else { [bool]($p -and $p.PSObject.Properties['value']) }
        if ($isList) { foreach ($v in @($p.value)) { if ($null -ne $v) { $out.Add($v) } } } elseif ($p) { $out.Add($p) }
        $uri = if ($p) { $p.'@odata.nextLink' } else { $null }
    }
    $out.ToArray()
}
function Select-FromList([string]$Title, [object[]]$Items, [scriptblock]$Label, [int]$Default = 0, [switch]$NoDefault) {
    # Numbered picker: type a number (or press Enter for the default). 0 stops. Works over any remote session.
    $Items = @($Items | Where-Object { $_ })
    if (-not $Items) { return $null }
    Write-Host "`n$Title" -ForegroundColor Cyan
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $mark = if (-not $NoDefault -and $i -eq $Default) { '   <- Enter' } else { '' }
        Write-Host ('  {0}  {1}{2}' -f ($i + 1), (& $Label $Items[$i]), $mark)
    }
    Write-Host '  0  Stop'
    while ($true) {
        $hint = if ($NoDefault -and $Items.Count -eq 1) { 'Type 1 to go ahead, 0 to stop' } elseif ($NoDefault) { "Number (1-$($Items.Count), 0 to stop)" } else { "Number (Enter = $($Default + 1), 0 to stop)" }
        $a = (Read-Host $hint).Trim()
        if ($a -eq '' -and -not $NoDefault) { return $Items[$Default] }
        if ($a -eq '0') { return $null }
        $n = 0
        if ([int]::TryParse($a, [ref]$n) -and $n -ge 1 -and $n -le $Items.Count) { return $Items[$n - 1] }
        Write-Host "Type a number from 0 to $($Items.Count)." -ForegroundColor Yellow
    }
}
function Confirm-Go([string]$Title, [string]$GoLabel = 'Yes, go ahead', [switch]$NoDefault) {
    # 1 = go, 0 = stop. -NoDefault for irreversible steps: Enter alone does nothing.
    $r = Select-FromList $Title @($GoLabel) { param($x) $x } -NoDefault:$NoDefault
    [bool]$r
}
function New-EasyPassword {
    # Easy to read out and type on a call (e.g. Move4821!). Meets Windows complexity rules. The account is
    # temporary: delete it once the move is done (or let LAPS manage it).
    'Move{0}!' -f (Get-Random -Minimum 1000 -Maximum 9999)
}
function Read-Choice($prompt, $options) {
    while ($true) { $a = Read-Host $prompt; if ($options -contains $a) { return $a }; Write-Host "Type one of: $($options -join ', ')" -ForegroundColor Yellow }
}
function Get-SiblingScript($name) {
    $local = Join-Path $PSScriptRoot $name
    if (Test-Path $local) { return $local }
    $dl = Join-Path $WorkFolder $name
    if (-not (Test-Path $dl)) { Invoke-WebRequest "$script:Repo/$name" -OutFile $dl -UseBasicParsing }
    $dl
}

# ---------------------------------------------------------------- 1. enrolment package (admin PC)
function Find-Wcd {
    # Windows Configuration Designer comes two ways: the Windows ADK (Program Files) or the Microsoft Store (WindowsApps)
    $adk = Get-ChildItem "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Imaging and Configuration Designer" -Recurse -Filter ICD.exe -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($adk) { return [pscustomobject]@{ Exe = $adk.FullName; Root = (Split-Path $adk.FullName); Source = 'ADK'; Launch = $adk.FullName } }
    $pkg = Get-AppxPackage -Name '*WindowsConfigurationDesigner*' -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1
    if ($pkg) {
        $exe = Join-Path $pkg.InstallLocation 'ICD.exe'
        $appId = try { @((Get-AppxPackageManifest $pkg).Package.Applications.Application)[0].Id } catch { 'App' }
        return [pscustomobject]@{ Exe = $(if (Test-Path $exe) { $exe } else { $null }); Root = $pkg.InstallLocation; Source = 'Store'; Launch = "shell:AppsFolder\$($pkg.PackageFamilyName)!$appId" }
    }
    $null
}
function Install-Wcd {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Warn2 'winget is not available. Install "Windows Configuration Designer" from the Microsoft Store, then run this option again.'
        Start-Process 'ms-windows-store://pdp/?ProductId=9NBLGGH4TX22' -ErrorAction SilentlyContinue
        return $null
    }
    Write-Host 'Windows Configuration Designer is not installed. Installing it from the Microsoft Store (winget)...' -ForegroundColor Cyan
    winget install --id 9NBLGGH4TX22 --source msstore --accept-package-agreements --accept-source-agreements --silent | Out-Host
    Start-Sleep 3
    $w = Find-Wcd
    if ($w) { Write-Ok "Windows Configuration Designer installed ($($w.Source) version)" } else { Write-Warn2 'Installed, but not found yet. Close and reopen PowerShell, then run this option again.' }
    $w
}

function Get-DrsAccessToken([string]$TenantId, [string]$Admin) {
    # Modern, interactive sign-in for the device registration service (OAuth 2.0 authorization code flow), so MFA
    # and Conditional Access apply and tenants that block device code still work. No libraries: plain PowerShell.
    # Runs in a clean child Windows PowerShell (STA) so it is unaffected by modules loaded in this session.
    # Tries: a sign-in window, then the system browser (localhost redirect with PKCE).
    # The Windows broker is not used: the first-party apps that can get this token have no broker redirect (AADSTS50011).
    $tmpDir = Join-Path $env:LOCALAPPDATA 'MigrationToolkit'; New-Item -ItemType Directory -Path $tmpDir -Force -WhatIf:$false | Out-Null
    $tokFile = Join-Path $tmpDir "drs-$([guid]::NewGuid()).tmp"
    $child = @'
$ErrorActionPreference = 'Stop'
$TenantId = '__TENANT__'; $Admin = '__ADMIN__'; $TokFile = '__TOKFILE__'
$tenantPath = if ($TenantId) { $TenantId } else { 'common' }
$drs = 'urn:ms-drs:enterpriseregistration.windows.net'
$aadps = '1b730954-1685-4b74-9bfd-dac224a7b894'   # Azure AD PowerShell: first-party client for device registration tokens
$wcd = 'b90d5b8f-5503-4153-b545-b31cecfaece2'     # Windows Configuration Designer
Add-Type -AssemblyName System.Windows.Forms, System.Web
function Enc([string]$v) { [uri]::EscapeDataString($v) }
function Get-AuthUrl($client, $redirect, $extra) {
    $u = "https://login.microsoftonline.com/$tenantPath/oauth2/authorize?client_id=$client&response_type=code&redirect_uri=$(Enc $redirect)&resource=$(Enc $drs)&prompt=select_account"
    if ($Admin) { $u += "&login_hint=$(Enc $Admin)" }
    $u + $extra
}
function Redeem($client, $code, $redirect, $verifier) {
    $body = @{ grant_type = 'authorization_code'; client_id = $client; code = $code; redirect_uri = $redirect; resource = $drs }
    if ($verifier) { $body.code_verifier = $verifier }
    $t = Invoke-RestMethod -Method Post "https://login.microsoftonline.com/$tenantPath/oauth2/token" -Body $body
    $p = $t.access_token.Split('.')[1].Replace('-', '+').Replace('_', '/'); while ($p.Length % 4) { $p += '=' }
    $c = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($p)) | ConvertFrom-Json
    $upn = if ($c.upn) { $c.upn } else { $c.unique_name }
    [IO.File]::WriteAllText($TokFile, (@{ Token = $t.access_token; TenantId = $c.tid; Upn = $upn } | ConvertTo-Json -Compress))
    Write-Host "   OK   Signed in as $upn" -ForegroundColor Green
}
function Read-Result([string]$url) {
    $q = [Web.HttpUtility]::ParseQueryString(([uri]$url).Query)
    if ($q['error']) { throw "$($q['error']): $($q['error_description'])" }
    $q['code']
}

# 1. Sign-in window (same approach as the Azure AD PowerShell module). Uses the modern rendering mode.
try {
    Write-Host 'Signing in with the sign-in window...' -ForegroundColor Cyan
    $fe = 'HKCU:\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION'
    if (-not (Test-Path $fe)) { New-Item $fe -Force | Out-Null }
    Set-ItemProperty $fe -Name 'powershell.exe' -Value 11001 -Type DWord
    $oob = 'urn:ietf:wg:oauth:2.0:oob'
    $form = New-Object Windows.Forms.Form -Property @{ Width = 520; Height = 720; Text = 'Sign in to the destination tenant'; StartPosition = 'CenterScreen'; TopMost = $true }
    $web = New-Object Windows.Forms.WebBrowser -Property @{ Dock = 'Fill'; ScriptErrorsSuppressed = $true }
    $form.Controls.Add($web)
    $script:result = $null
    $check = { param($u) if ($u -match 'oauth2/nativeclient' -or $u -like 'urn:ietf:wg:oauth*') { $script:result = $u; $form.Close() } }
    $web.add_Navigating({ param($s, $e) & $check $e.Url.AbsoluteUri })
    $web.add_Navigated({ param($s, $e) & $check $e.Url.AbsoluteUri })
    $web.Navigate((Get-AuthUrl $aadps $oob ''))
    [void]$form.ShowDialog()
    if (-not $script:result) { throw 'the window was closed before sign-in finished' }
    Redeem $aadps (Read-Result $script:result) $oob $null
    exit 0
}
catch { Write-Host "   WARN sign-in window: $($_.Exception.Message)" -ForegroundColor Yellow }

# 2. System browser with a localhost redirect and PKCE
foreach ($client in $aadps, $wcd) {
    $http = $null
    try {
        Write-Host "Signing in with the system browser ($(if ($client -eq $aadps) { 'device registration client' } else { 'configuration designer client' }))..." -ForegroundColor Cyan
        $l = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, 0); $l.Start(); $port = $l.LocalEndpoint.Port; $l.Stop()
        $redirect = "http://localhost:$port/"
        $bytes = New-Object byte[] 32; [Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($bytes)
        $verifier = [Convert]::ToBase64String($bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        $hash = [Security.Cryptography.SHA256]::Create().ComputeHash([Text.Encoding]::ASCII.GetBytes($verifier))
        $challenge = [Convert]::ToBase64String($hash).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        $http = [Net.HttpListener]::new(); $http.Prefixes.Add($redirect); $http.Start()
        Start-Process (Get-AuthUrl $client $redirect "&code_challenge=$challenge&code_challenge_method=S256")
        Write-Host '   Finish the sign-in in your browser (this attempt gives up after 3 minutes; if the browser shows an error, wait for the next one).' -ForegroundColor Gray
        $task = $http.GetContextAsync()
        if (-not $task.Wait([TimeSpan]::FromMinutes(3))) { throw 'timed out waiting for the browser' }
        $ctx = $task.Result
        $msg = [Text.Encoding]::UTF8.GetBytes('<html><body style="font-family:Segoe UI">Signed in. You can close this tab and go back to PowerShell.</body></html>')
        $ctx.Response.ContentType = 'text/html'; $ctx.Response.OutputStream.Write($msg, 0, $msg.Length); $ctx.Response.Close()
        Redeem $client (Read-Result $ctx.Request.Url.AbsoluteUri) $redirect $verifier
        exit 0
    }
    catch { Write-Host "   WARN system browser: $($_.Exception.Message)" -ForegroundColor Yellow }
    finally { if ($http) { $http.Close() } }
}
exit 1
'@
    $child = $child.Replace('__TENANT__', $TenantId).Replace('__ADMIN__', ($Admin -replace "'", "''")).Replace('__TOKFILE__', ($tokFile -replace "'", "''"))
    $enc = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($child))
    $exe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    try {
        & $exe -NoProfile -STA -ExecutionPolicy Bypass -EncodedCommand $enc | Out-Host
        if (Test-Path $tokFile) { return (Get-Content $tokFile -Raw | ConvertFrom-Json) }
    }
    finally { Remove-Item $tokFile -Force -ErrorAction SilentlyContinue }
    throw 'No interactive sign-in method was accepted (see the warnings above).'
}
function Get-JwtClaims([string]$Jwt) {
    $p = $Jwt.Split('.')[1].Replace('-', '+').Replace('_', '/'); while ($p.Length % 4) { $p += '=' }
    [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($p)) | ConvertFrom-Json
}
function Get-DrsAccessTokenDeviceCode([string]$TenantId, [string]$Admin) {
    # Only with -UseDeviceCode: many tenants block device code flow by Conditional Access
    if (-not $TenantId) { $TenantId = 'organizations' }
    $client = 'b90d5b8f-5503-4153-b545-b31cecfaece2'; $drs = 'urn:ms-drs:enterpriseregistration.windows.net'
    $dc = Invoke-RestMethod -Method Post "https://login.microsoftonline.com/$TenantId/oauth2/devicecode" -Body @{ client_id = $client; resource = $drs }
    Write-Host $dc.message -ForegroundColor Yellow
    $deadline = (Get-Date).AddSeconds([int]$dc.expires_in)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds ([int]$dc.interval)
        try { $t = Invoke-RestMethod -Method Post "https://login.microsoftonline.com/$TenantId/oauth2/token" -Body @{ grant_type = 'device_code'; client_id = $client; code = $dc.device_code; resource = $drs }; if ($t.access_token) { $c = Get-JwtClaims $t.access_token; return [pscustomobject]@{ Token = $t.access_token; TenantId = $c.tid; Upn = $c.upn } } }
        catch { $bd = ''; try { $bd = $_.ErrorDetails.Message } catch {}; if ($bd -notmatch 'authorization_pending') { throw } }
    }
    throw 'Sign-in timed out.'
}
function Get-BulkToken($SignIn, [int]$Days) {
    # The same call Windows Configuration Designer makes behind "Get Bulk Token": it creates the package_<guid>
    # account in the tenant and returns a bulk token that lets devices join without a user signing in.
    $h = @{ Authorization = "Bearer $($SignIn.Token)" }
    $pid0 = [guid]::NewGuid().ToString()
    $body = @{ pid = $pid0; name = "package_$pid0"; exp = (Get-Date).AddDays($Days).ToString('MM/dd/yyyy') } | ConvertTo-Json
    $r = Invoke-RestMethod -Method Post 'https://login.microsoftonline.com/webapp/bulkaadjtoken/begin' -Headers $h -Body $body -ContentType 'application/json'
    $until = (Get-Date).AddMinutes(3)
    while ($r.state -notmatch 'Complete|Fail|Error' -and (Get-Date) -lt $until) {
        Start-Sleep -Seconds 3
        $r = Invoke-RestMethod -Method Get "https://login.microsoftonline.com/webapp/bulkaadjtoken/poll?flowToken=$($r.flowToken)" -Headers $h
    }
    if ($r.state -ne 'CompleteSuccess') { throw "Bulk token request ended with state '$($r.state)' $($r.resultData)" }
    $data = $r.resultData; if ($data -is [string]) { try { $data = $data | ConvertFrom-Json } catch {} }
    foreach ($k in 'rt', 'BPRT', 'bprt', 'refresh_token', 'token') { if ($data -and $data.PSObject.Properties[$k] -and $data.$k) { return $data.$k } }
    if ($r.resultData -is [string] -and $r.resultData.Length -gt 200) { return $r.resultData }
    throw 'Bulk token request succeeded but no token was found in the response.'
}
function Get-EventSourceDll {
    # ADK 10.1.26100+ ships ICD.exe without Microsoft.Diagnostics.Tracing.EventSource.dll, so it crashes on start.
    # Take it from the Store designer if present, otherwise from its NuGet package (version 1.1.28).
    $cache = Join-Path $env:LOCALAPPDATA 'MigrationToolkit\EventSource'
    $dll = Get-ChildItem $cache -Recurse -Filter 'Microsoft.Diagnostics.Tracing.EventSource.dll' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($dll) { return $dll.FullName }
    $dll = Get-ChildItem (Join-Path $env:LOCALAPPDATA 'MigrationToolkit') -Recurse -Filter 'Microsoft.Diagnostics.Tracing.EventSource.dll' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($dll) { return $dll.FullName }
    try {
        New-Item -ItemType Directory -Path $cache -Force -WhatIf:$false | Out-Null
        $zip = Join-Path $cache 'eventsource.zip'
        Invoke-WebRequest 'https://www.nuget.org/api/v2/package/Microsoft.Diagnostics.Tracing.EventSource.Redist/1.1.28' -OutFile $zip -UseBasicParsing
        Expand-Archive $zip -DestinationPath $cache -Force
        $dll = Get-ChildItem $cache -Recurse -Filter 'Microsoft.Diagnostics.Tracing.EventSource.dll' | Where-Object { $_.FullName -match 'net4' } | Select-Object -First 1
        if ($dll) { return $dll.FullName }
    } catch { Write-Warn2 "Could not download the EventSource library: $($_.Exception.Message)" }
    $null
}
function Copy-IcdLocal([string]$From, [string]$Name) {
    # Private copy under LocalAppData: runs without admin and can be patched without touching Program Files
    $dest = Join-Path $env:LOCALAPPDATA "MigrationToolkit\$Name"
    if (-not (Test-Path (Join-Path $dest 'ICD.exe'))) {
        New-Item -ItemType Directory -Path $dest -Force -WhatIf:$false | Out-Null
        robocopy $From $dest /E /R:1 /W:1 /NFL /NDL /NJH /NJS /NP | Out-Null
    }
    if (-not (Test-Path (Join-Path $dest 'ICD.exe'))) { return $null }
    if (-not (Test-Path (Join-Path $dest 'Microsoft.Diagnostics.Tracing.EventSource.dll'))) {
        $es = Get-EventSourceDll
        if ($es) { Copy-Item $es $dest -Force -WhatIf:$false }
    }
    $dest
}
function Get-IcdCandidates([object]$Wcd) {
    # Command-line builders. Both are run from a private copy: the Store app can't be driven from the command line
    # inside its sandbox, and the current ADK build is missing a library it needs.
    $list = @()
    $pkg = Get-AppxPackage -Name '*WindowsConfigurationDesigner*' -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1
    if ($pkg) {
        $exeDir = Get-ChildItem $pkg.InstallLocation -Recurse -Filter ICD.exe -ErrorAction SilentlyContinue | Select-Object -First 1
        $d = if ($exeDir) { Copy-IcdLocal $exeDir.DirectoryName "ICD-Store-$($pkg.Version)" }
        if ($d) { $list += [pscustomobject]@{ Exe = (Join-Path $d 'ICD.exe'); Root = $d; Source = 'Store (copied out)'; Launch = $null } }
    }
    $adk = Find-Wcd | Where-Object Source -eq 'ADK'
    if ($adk) {
        $d = Copy-IcdLocal $adk.Root 'ICD-ADK'
        if ($d) { $list += [pscustomobject]@{ Exe = (Join-Path $d 'ICD.exe'); Root = $d; Source = 'ADK (copied out)'; Launch = $null } }
    }
    $list
}
function Invoke-Icd($Wcd, [string[]]$IcdArgs, [string]$Log) {
    $out = "$Log.out"; $err = "$Log.err"
    $p = Start-Process -FilePath $Wcd.Exe -ArgumentList $IcdArgs -WorkingDirectory (Split-Path $Wcd.Exe) -NoNewWindow -Wait -PassThru -RedirectStandardOutput $out -RedirectStandardError $err
    $text = ((Get-Content $out, $err -Raw -ErrorAction SilentlyContinue) -join "`n").Trim()
    Add-Content $Log "[$($Wcd.Source)] $($Wcd.Exe) exit $($p.ExitCode)`n$text`n" -WhatIf:$false
    Remove-Item $out, $err -Force -ErrorAction SilentlyContinue
    if ($p.ExitCode -ne 0) { throw "ICD exit code $($p.ExitCode). $(($text -split "`n" | Where-Object { $_ -match '\S' } | Select-Object -Last 3) -join ' | ')" }
}
function Install-AdkIcd {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) { Write-Warn2 'winget is not available, so the ADK designer cannot be installed automatically.'; return $null }
    Write-Host 'Installing the Windows ADK Configuration Designer only (silent, about 5 minutes; Windows may ask once to allow it)...' -ForegroundColor Cyan
    winget install --id Microsoft.WindowsADK --accept-package-agreements --accept-source-agreements --silent --override '/features OptionId.ImagingAndConfigurationDesigner OptionId.ICDConfigurationDesigner /quiet /norestart /ceip off' | Out-Host
    $w = Find-Wcd | Where-Object Source -eq 'ADK'
    if ($w) { Write-Ok 'Configuration Designer (ADK) installed'; return $w }
    Write-Warn2 'ADK install finished but ICD.exe was not found.'
    $null
}

function Invoke-BuildPackage {
    Write-Step 'Build the enrolment package (run on your admin PC)'
    $wcd = Find-Wcd
    if ($wcd) { Write-Ok "Windows Configuration Designer found ($($wcd.Source) version)" }

    # Sign in first: the tenant, the account and a sensible package name all come from the sign-in
    $signIn = $null
    try { $signIn = if ($UseDeviceCode) { Get-DrsAccessTokenDeviceCode $DestinationTenantId $DestinationAdmin } else { Get-DrsAccessToken $DestinationTenantId $DestinationAdmin } }
    catch { Write-Warn2 "Sign-in failed: $($_.Exception.Message)"; Write-Warn2 'The package was not built. Nothing was opened.'; return }
    $domain = if ($signIn.Upn -match '@(.+)$') { $Matches[1] } else { 'tenant' }
    Write-Ok "Destination tenant: $domain ($($signIn.TenantId))"
    $name = if ($PackageName) { $PackageName } else { "$(($domain -split '\.')[0])-join-$(Get-Date -Format yyyyMMdd-HHmm)" }
    $name = $name -replace '[^A-Za-z0-9-]', ''
    $expiry = $TokenDays
    $rename = $RenamePattern
    if ($rename -and ($rename -replace '%SERIAL%', 'XXXXXXXX' -replace '%RAND:(\d+)%', 'XXXX').Length -gt 15) { Write-Warn2 "Rename pattern '$rename' may exceed 15 characters on some laptops; Windows truncates it." }
    $ppkg = Join-Path $WorkFolder "$name.ppkg"
    $log = Join-Path $WorkFolder "ICD-build-$script:Stamp.log"

    $built = $false
    $bprt = $null
    try { $bprt = Get-BulkToken $signIn $expiry; Write-Ok "Bulk token received (valid $expiry days)" }
    catch { Write-Warn2 "Bulk token request failed: $($_.Exception.Message)" }

    if ($bprt) {
        $id = [guid]::NewGuid().ToString()
        $renameXml = if ($rename) { "<DevDetail><DNSComputerName>$rename</DNSComputerName></DevDetail>" } else { '' }
        $xml = @"
<?xml version="1.0" encoding="utf-8"?>
<WindowsCustomizations>
  <PackageConfig xmlns="urn:schemas-Microsoft-com:Windows-ICD-Package-Config.v1.0">
    <ID>{$id}</ID><Name>$name</Name><Version>1.0</Version><OwnerType>OEM</OwnerType><Rank>0</Rank>
  </PackageConfig>
  <Settings xmlns="urn:schemas-microsoft-com:windows-provisioning">
    <Customizations><Common>
      <Accounts><Azure><Authority>https://login.microsoftonline.com/common</Authority><BPRT>$bprt</BPRT></Azure></Accounts>
      $renameXml
    </Common></Customizations>
  </Settings>
</WindowsCustomizations>
"@
        $xmlPath = Join-Path $WorkFolder "$name.customizations.xml"
        $xml | Set-Content $xmlPath -Encoding UTF8 -WhatIf:$false
        $bprt = $null
        $tried = @()
        $cands = @(Get-IcdCandidates $wcd)
        foreach ($w in $cands + @('ADK')) {
            if ($w -eq 'ADK') {
                if ($tried -like 'ADK*') { break }
                $w = Install-AdkIcd; if (-not $w) { break }
                $d = Copy-IcdLocal $w.Root 'ICD-ADK'; if (-not $d) { break }
                $w = [pscustomobject]@{ Exe = (Join-Path $d 'ICD.exe'); Root = $d; Source = 'ADK (copied out)'; Launch = $null }
            }
            $tried += $w.Source
            $storeFile = Get-ChildItem $w.Root -Recurse -Filter 'Microsoft-Desktop-Provisioning.dat' -ErrorAction SilentlyContinue | Select-Object -First 1
            $icdArgs = @('/Build-ProvisioningPackage', "/CustomizationXML:`"$xmlPath`"", "/PackagePath:`"$ppkg`"", '+Overwrite')
            if ($storeFile) { $icdArgs += "/StoreFile:`"$($storeFile.FullName)`"" }
            Write-Host "Building with the $($w.Source) Configuration Designer..." -ForegroundColor Cyan
            try { Invoke-Icd $w $icdArgs $log } catch { Write-Warn2 "$($w.Source) build failed: $($_.Exception.Message)" }
            if (Test-Path $ppkg) { $built = $true; break }
        }
        Remove-Item $xmlPath -Force -ErrorAction SilentlyContinue   # it holds the token
    }

    if ($built) {
        Write-Ok "Package built: $ppkg"
        Write-Host 'Keep it private: it joins any device to the destination tenant until the token expires. Delete it when the move is done.' -ForegroundColor Yellow
        Add-State 'Package' $ppkg
        return
    }

    if (-not $AllowWizard) {
        Write-Warn2 "The package was not built. Nothing was opened. Build log: $log"
        Write-Host 'Paste the warning lines above (and the log) to fix it. Add -AllowWizard to fall back to the Windows Configuration Designer wizard.' -ForegroundColor Yellow
        return
    }
    if (-not $wcd) { $wcd = Install-Wcd; if (-not $wcd) { return } }
    # Guided: Windows Configuration Designer's own wizard (requests its own token)
    Write-Step 'Guided: build it in Windows Configuration Designer (about 3 minutes)'
    Write-Host @"
1. Click "Provision desktop devices".
2. Name: $name. Next.
3. Set up device: leave the device name BLANK (a pattern renames every laptop and adds a restart). Turn off "Remove pre-installed software". Next.
4. Set up network: turn it OFF. Next.
5. Account management: choose "Enroll in Azure AD". Bulk Token Expiry: $expiry days from today. Click "Get Bulk Token" and sign in as $($signIn.Upn).
   If it asks "Allow my organization to manage my device", untick it and choose "This app only".
6. Add applications / certificates: skip. Finish: click Create. It shows the folder it saved to.
"@ -ForegroundColor Yellow
    if ($wcd.Source -eq 'Store') { Start-Process 'explorer.exe' $wcd.Launch -ErrorAction SilentlyContinue } else { Start-Process $wcd.Launch -ErrorAction SilentlyContinue }
    $started = Get-Date
    $p = Read-Host 'Wait for "Your provisioning package has been successfully built", then press Enter (or paste the .ppkg path)'
    if (-not $p) {
        # WCD's save folder varies by version and by OneDrive folder redirection, so search instead of guessing
        $roots = @([Environment]::GetFolderPath('MyDocuments'), $env:USERPROFILE, (Join-Path $env:LOCALAPPDATA 'Packages')) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique
        $hit = foreach ($r in $roots) { Get-ChildItem $r -Recurse -Filter '*.ppkg' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -gt $started.AddHours(-2) } }
        $hit = @($hit | Sort-Object LastWriteTime -Descending | Select-Object -Unique)
        $exact = @($hit | Where-Object { $_.BaseName -eq $name })
        $p = if ($exact) { $exact[0].FullName } elseif ($hit.Count -eq 1) { $hit[0].FullName } elseif ($hit) { $hit | Select-Object FullName, LastWriteTime | Format-Table -AutoSize | Out-String -Width 250 | Write-Host; Read-Host 'Several packages found. Paste the one to use' } else { '' }
        if ($p) { Write-Host "Found: $p" -ForegroundColor Cyan }
    }
    if ($p -and (Test-Path $p)) { Copy-Item $p $ppkg -Force -WhatIf:$false; Write-Ok "Package ready: $ppkg"; Add-State 'Package' $ppkg }
    else { Write-Warn2 "No new .ppkg found yet. Finish the wizard (Create), then run option 1 again or pass the .ppkg to option 4 with -PackagePath." }
}

# ---------------------------------------------------------------- 2. Intune configuration (admin PC)
$script:IntuneTypes = [ordered]@{
    'deviceConfigurations'               = @{ Uri = 'deviceManagement/deviceConfigurations'; Name = 'displayName'; Assign = 'assign'; AssignKey = 'assignments' }
    'configurationPolicies'              = @{ Uri = 'deviceManagement/configurationPolicies'; Name = 'name'; Assign = 'assign'; AssignKey = 'assignments'; Settings = $true }
    'deviceCompliancePolicies'           = @{ Uri = 'deviceManagement/deviceCompliancePolicies'; Name = 'displayName'; Assign = 'assign'; AssignKey = 'assignments'; Compliance = $true }
    'deviceManagementScripts'            = @{ Uri = 'deviceManagement/deviceManagementScripts'; Name = 'displayName'; Assign = 'assign'; AssignKey = 'deviceManagementScriptAssignments'; Content = $true }
    'deviceHealthScripts'                = @{ Uri = 'deviceManagement/deviceHealthScripts'; Name = 'displayName'; Assign = 'assign'; AssignKey = 'deviceHealthScriptAssignments'; Content = $true }
    'windowsAutopilotDeploymentProfiles' = @{ Uri = 'deviceManagement/windowsAutopilotDeploymentProfiles'; Name = 'displayName'; Assign = $null }
    'mobileApps'                         = @{ Uri = 'deviceAppManagement/mobileApps'; Name = 'displayName'; Assign = $null; MetadataOnly = $true }
}
$script:ReadScopes = 'DeviceManagementConfiguration.Read.All', 'DeviceManagementScripts.Read.All', 'DeviceManagementApps.Read.All', 'DeviceManagementServiceConfig.Read.All', 'Group.Read.All'
$script:WriteScopes = 'DeviceManagementConfiguration.ReadWrite.All', 'DeviceManagementScripts.ReadWrite.All', 'DeviceManagementServiceConfig.ReadWrite.All', 'Group.Read.All'
$script:StripProps = 'id', 'createdDateTime', 'lastModifiedDateTime', 'version', 'supportsScopeTags', 'isAssigned', 'settingCount', 'priorityMetaData', 'assignments', 'deviceManagementScriptAssignments', 'deviceHealthScriptAssignments', 'runSummary', 'deviceRunStates', 'userRunStates', 'groupAssignments', 'highestAvailableVersion', 'isGlobalScript', 'managedDeviceCount', 'uploadState', 'publishingState', 'isFeatured', 'size', 'dependentAppCount', 'supersedingAppCount', 'supersededAppCount', 'committedContentVersion', 'contentVersions', 'installSummary', 'relationships', 'categories'

function Invoke-CopyIntune {
    Write-Step 'Copy Intune configuration'
    $mode = Read-Choice 'Export from the source (E), import into the destination (I), or both (B)?' @('E', 'I', 'B', 'e', 'i', 'b')
    $mode = $mode.ToUpper()
    if (-not $ExportFolder) { $ExportFolder = Join-Path $WorkFolder "IntuneExport-$script:Stamp" }

    if ($mode -in 'E', 'B') {
        if (-not $SourceTenantId) { $SourceTenantId = Read-Host 'Source tenant ID' }
        if (-not $SourceAdmin) { $SourceAdmin = Read-Host 'Source admin sign-in' }
        Connect-Tenant $SourceTenantId $SourceAdmin $script:ReadScopes 'SOURCE'
        New-Item -ItemType Directory -Path $ExportFolder -Force -WhatIf:$false | Out-Null
        $groups = @{}
        Get-GraphAll 'v1.0/groups?$select=id,displayName&$top=999' | ForEach-Object { $groups[$_.id] = $_.displayName }
        foreach ($t in $script:IntuneTypes.Keys) {
            $d = $script:IntuneTypes[$t]
            $items = @(Get-GraphAll "beta/$($d.Uri)?`$top=100")
            $rows = foreach ($i in $items) {
                $o = @{}; $i.Keys | ForEach-Object { $o[$_] = $i[$_] }
                if ($d.Settings) { $o['settings'] = @(Get-GraphAll "beta/$($d.Uri)/$($i.id)/settings") }
                if ($d.Content -and -not $i.scriptContent) { $full = Invoke-MgGraphRequest GET "beta/$($d.Uri)/$($i.id)"; $o['scriptContent'] = $full.scriptContent; if ($full.detectionScriptContent) { $o['detectionScriptContent'] = $full.detectionScriptContent; $o['remediationScriptContent'] = $full.remediationScriptContent } }
                if ($d.Compliance) { $o['scheduledActionsForRule'] = @(Get-GraphAll "beta/$($d.Uri)/$($i.id)/scheduledActionsForRule?`$expand=scheduledActionConfigurations") }
                if ($d.Assign) {
                    $a = @(Get-GraphAll "beta/$($d.Uri)/$($i.id)/$($d.AssignKey)")
                    $o['_assignments'] = @($a | ForEach-Object { $tg = $_.target; [pscustomobject]@{ type = $tg.'@odata.type'; groupId = $tg.groupId; groupName = $(if ($tg.groupId) { $groups[$tg.groupId] }); filterId = $tg.deviceAndAppManagementAssignmentFilterId; filterType = $tg.deviceAndAppManagementAssignmentFilterType; intent = $_.intent } })
                }
                $o
            }
            $rows | ConvertTo-Json -Depth 30 | Set-Content (Join-Path $ExportFolder "$t.json") -Encoding UTF8 -WhatIf:$false
            Write-Ok "$t : $($items.Count)"
        }
        $groups.GetEnumerator() | ForEach-Object { [pscustomobject]@{ SourceGroupId = $_.Key; GroupName = $_.Value; DestinationGroupName = $_.Value } } | Export-Csv (Join-Path $ExportFolder 'GroupMapping.csv') -NoTypeInformation -WhatIf:$false
        Write-Host "Exported to $ExportFolder. Edit GroupMapping.csv if destination groups have different names." -ForegroundColor Cyan
        Write-Host 'Apps: only metadata is exported. Win32 and LOB installers must be re-uploaded in the destination; mobileApps.json is your list.' -ForegroundColor Yellow
    }

    if ($mode -in 'I', 'B') {
        if (-not (Test-Path $ExportFolder)) { throw "Export folder $ExportFolder not found." }
        if (-not $DestinationTenantId) { $DestinationTenantId = Read-Host 'Destination tenant ID' }
        if (-not $DestinationAdmin) { $DestinationAdmin = Read-Host 'Destination admin sign-in' }
        Connect-Tenant $DestinationTenantId $DestinationAdmin $script:WriteScopes 'DESTINATION'
        $dstGroups = @{}
        Get-GraphAll 'v1.0/groups?$select=id,displayName&$top=999' | ForEach-Object { $dstGroups[$_.displayName] = $_.id }
        $map = @{}
        $mf = if ($GroupMappingFile) { $GroupMappingFile } else { Join-Path $ExportFolder 'GroupMapping.csv' }
        if (Test-Path $mf) { Import-Csv $mf | ForEach-Object { if ($_.DestinationGroupName) { $map[$_.SourceGroupId] = $_.DestinationGroupName } } }
        $report = [System.Collections.Generic.List[object]]::new()
        foreach ($t in $script:IntuneTypes.Keys) {
            $d = $script:IntuneTypes[$t]
            $f = Join-Path $ExportFolder "$t.json"
            if (-not (Test-Path $f)) { continue }
            $items = @(ConvertFrom-Json (Get-Content $f -Raw))
            if (-not $items) { continue }
            if ($d.MetadataOnly) { $items | ForEach-Object { $report.Add([pscustomobject]@{ Type = $t; Name = $_.displayName; Status = 'Listed only'; Detail = "$($_.'@odata.type') must be re-uploaded by hand" }) }; continue }
            $existing = @{}
            Get-GraphAll "beta/$($d.Uri)?`$top=100" | ForEach-Object { $k = "$($_[$d.Name])"; if ($k) { $existing[$k] = $_.id } }
            foreach ($i in $items) {
                $nm = "$($i.($d.Name))"
                if (-not $nm) { $nm = "$($i.displayName)$($i.name)" }
                if ($nm -and $existing.ContainsKey($nm)) { $report.Add([pscustomobject]@{ Type = $t; Name = $nm; Status = 'Exists'; Detail = 'Same name already in the destination, not changed' }); continue }
                $body = @{}
                $i.PSObject.Properties | Where-Object { $_.Name -notin $script:StripProps -and $_.Name -notlike '_*' -and $_.Name -notlike '*@odata.context' } | ForEach-Object { $body[$_.Name] = $_.Value }
                if ($d.Compliance -and -not $body.scheduledActionsForRule) { $body.scheduledActionsForRule = @(@{ ruleName = 'PasswordRequired'; scheduledActionConfigurations = @(@{ actionType = 'block'; gracePeriodHours = 0; notificationTemplateId = ''; notificationMessageCCList = @() }) }) }
                if ($d.Compliance -and $body.scheduledActionsForRule) { $body.scheduledActionsForRule = @($body.scheduledActionsForRule | ForEach-Object { $r = $_; @{ ruleName = $r.ruleName; scheduledActionConfigurations = @($r.scheduledActionConfigurations | ForEach-Object { @{ actionType = $_.actionType; gracePeriodHours = $_.gracePeriodHours; notificationTemplateId = ''; notificationMessageCCList = @() } }) } }) }
                if (-not $PSCmdlet.ShouldProcess("$t '$nm'", 'Create in the destination')) { $report.Add([pscustomobject]@{ Type = $t; Name = $nm; Status = 'WhatIf'; Detail = "Would create, then assign to: $(($i._assignments | ForEach-Object { if ($_.groupName) { $_.groupName } else { $_.type -replace '#microsoft.graph.', '' } }) -join ', ')" }); continue }
                try {
                    $new = Invoke-MgGraphRequest POST "beta/$($d.Uri)" -ContentType 'application/json' -Body ($body | ConvertTo-Json -Depth 30)
                    $detail = 'Created'
                    if ($d.Assign -and $i._assignments) {
                        $targets = foreach ($a in $i._assignments) {
                            $tg = @{ '@odata.type' = $a.type }
                            if ($a.groupId) {
                                $dn = if ($map[$a.groupId]) { $map[$a.groupId] } else { $a.groupName }
                                if ($dn -and $dstGroups[$dn]) { $tg.groupId = $dstGroups[$dn] } else { $detail += "; group '$($a.groupName)' not found in destination, assignment skipped"; continue }
                            }
                            $as = @{ target = $tg }
                            if ($a.intent) { $as.intent = $a.intent }
                            $as
                        }
                        if ($targets) {
                            $k = if ($t -eq 'deviceManagementScripts') { 'deviceManagementScriptAssignments' } elseif ($t -eq 'deviceHealthScripts') { 'deviceHealthScriptAssignments' } else { 'assignments' }
                            Invoke-MgGraphRequest POST "beta/$($d.Uri)/$($new.id)/assign" -ContentType 'application/json' -Body (@{ $k = @($targets) } | ConvertTo-Json -Depth 10) | Out-Null
                            $detail += "; assigned ($(@($targets).Count))"
                        }
                    }
                    $report.Add([pscustomobject]@{ Type = $t; Name = $nm; Status = 'Done'; Detail = $detail })
                }
                catch { $report.Add([pscustomobject]@{ Type = $t; Name = $nm; Status = 'Problem'; Detail = $_.Exception.Message }) }
            }
        }
        $rf = Join-Path $WorkFolder "IntuneImport-$script:Stamp.csv"
        $report | Export-Csv $rf -NoTypeInformation -WhatIf:$false
        $report | Format-Table Type, Name, Status, Detail -AutoSize -Wrap | Out-String -Width 250 | Write-Host
        Write-Host "Report: $rf" -ForegroundColor Cyan
    }
}

# ---------------------------------------------------------------- 3. Autopilot
function Invoke-AutopilotCapture {
    Write-Step 'Capture this device''s Autopilot hardware hash'
    $dev = Get-CimInstance -Namespace root/cimv2/mdm/dmmap -ClassName MDM_DevDetail_Ext01 -Filter "InstanceID='Ext' AND ParentID='./DevDetail'" -ErrorAction SilentlyContinue
    $hash = if ($dev) { $dev.DeviceHardwareData } else { $null }
    $serial = (Get-CimInstance Win32_BIOS).SerialNumber
    if (-not $hash) { Write-Warn2 'No hardware hash available (needs an elevated session on a physical Windows device).'; return }
    $f = Join-Path $WorkFolder "AutopilotHash-$env:COMPUTERNAME.csv"
    [pscustomobject]@{ 'Device Serial Number' = $serial; 'Windows Product ID' = ''; 'Hardware Hash' = $hash } | Export-Csv $f -NoTypeInformation -WhatIf:$false
    Write-Ok "Saved $f (serial $serial). Copy it to your admin PC for option 3 > move."
    Add-State 'AutopilotHashFile' $f
}
function Invoke-AutopilotMove {
    Write-Step 'Register Autopilot hashes in the destination and remove them from the source'
    $folder = Read-Host "Folder holding AutopilotHash-*.csv files (default $WorkFolder)"
    if (-not $folder) { $folder = $WorkFolder }
    $files = @(Get-ChildItem $folder -Filter 'AutopilotHash-*.csv' -Recurse)
    if (-not $files) { Write-Warn2 'No AutopilotHash-*.csv files found.'; return }
    $rows = $files | ForEach-Object { Import-Csv $_.FullName }
    if (-not $DestinationTenantId) { $DestinationTenantId = Read-Host 'Destination tenant ID' }
    if (-not $DestinationAdmin) { $DestinationAdmin = Read-Host 'Destination admin sign-in' }
    Connect-Tenant $DestinationTenantId $DestinationAdmin @('DeviceManagementServiceConfig.ReadWrite.All') 'DESTINATION'
    $tag = Read-Host 'Group tag for these devices in the destination (optional)'
    foreach ($r in $rows) {
        if ($PSCmdlet.ShouldProcess("Serial $($r.'Device Serial Number')", 'Import into destination Autopilot')) {
            $b = @{ '@odata.type' = '#microsoft.graph.importedWindowsAutopilotDeviceIdentity'; serialNumber = $r.'Device Serial Number'; hardwareIdentifier = $r.'Hardware Hash'; groupTag = $tag; state = @{ '@odata.type' = 'microsoft.graph.importedWindowsAutopilotDeviceIdentityState'; deviceImportStatus = 'pending'; deviceRegistrationId = ''; deviceErrorCode = 0; deviceErrorName = '' } }
            Invoke-MgGraphRequest POST 'beta/deviceManagement/importedWindowsAutopilotDeviceIdentities' -ContentType 'application/json' -Body ($b | ConvertTo-Json -Depth 5) | Out-Null
            Write-Ok "Imported $($r.'Device Serial Number') (Intune processes imports in the background, allow 15 minutes)"
        }
    }
    $rm = Read-Choice 'Remove the same serials from the SOURCE tenant''s Autopilot now? (Y/N)' @('Y', 'N', 'y', 'n')
    if ($rm -in 'Y', 'y') {
        if (-not $SourceTenantId) { $SourceTenantId = Read-Host 'Source tenant ID' }
        if (-not $SourceAdmin) { $SourceAdmin = Read-Host 'Source admin sign-in' }
        Connect-Tenant $SourceTenantId $SourceAdmin @('DeviceManagementServiceConfig.ReadWrite.All') 'SOURCE'
        foreach ($r in $rows) {
            $s = $r.'Device Serial Number'
            $hit = @(Get-GraphAll "beta/deviceManagement/windowsAutopilotDeviceIdentities?`$filter=contains(serialNumber,'$s')")
            if (-not $hit) { Write-Warn2 "$s not registered in the source"; continue }
            foreach ($h in $hit) { if ($PSCmdlet.ShouldProcess("Serial $s", 'Remove from source Autopilot')) { Invoke-MgGraphRequest DELETE "beta/deviceManagement/windowsAutopilotDeviceIdentities/$($h.id)" | Out-Null; Write-Ok "Removed $s from the source" } }
        }
    }
}

# ---------------------------------------------------------------- 4. move this laptop (staged)
function Invoke-MoveDevice {
    if (-not (Test-Elevated)) { throw 'Run as administrator on the laptop.' }
    $state = Get-State
    $ds = Get-DsregStatus
    Write-Step "Move this laptop: stage so far = $($state.Stage)"

    if ($state.Stage -in 'NotStarted', 'PreflightFailed') {
        # ----- Stage 1: pre-flight and safety net
        if ($ds.DomainJoined -eq 'YES') { throw 'This PC is joined to an on-premises domain. This script is for Microsoft Entra joined devices only.' }
        if ($ds.AzureAdJoined -ne 'YES') { throw 'This PC is not Microsoft Entra joined. Nothing to leave.' }
        Write-Ok "Entra joined to $($ds.TenantName) ($($ds.TenantId))"
        if ($SourceTenantId) {
            if ($ds.TenantId.Trim('{} ') -ne $SourceTenantId.Trim('{} ')) {
                throw "STOP: this machine is joined to $($ds.TenantName) ($($ds.TenantId)), not the source tenant $SourceTenantId. Nothing has changed. Option 4 only runs on a laptop being moved out of the source tenant; never on your own admin PC."
            }
            Write-Ok 'Joined tenant matches -SourceTenantId'
        }
        else {
            Write-Warn2 "No -SourceTenantId given. Before leaving you will have to type the tenant name '$($ds.TenantName)' exactly. If this is your own PC, stop now."
        }
        $free = [math]::Round((Get-PSDrive C).Free / 1GB, 1); if ($free -lt 5) { Write-Warn2 "Only $free GB free on C:" } else { Write-Ok "$free GB free on C:" }
        if ($env:USERNAME -and (Get-Profiles | Where-Object { $_.Loaded -and $_.IsEntra })) { Write-Warn2 'A Microsoft Entra user is signed in. Finish this stage, then restart and sign in as the local admin for stage 2.' }

        $profiles = @(Get-Profiles)
        if ($OldProfile) { $old = $profiles | Where-Object { $_.Folder -eq (Split-Path $OldProfile -Leaf) } | Select-Object -First 1 }
        else {
            $me = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
            $cands = @($profiles | Where-Object { $_.Sid -ne $me -and $_.Folder -ne $LocalAdminName } | Sort-Object @{ e = { $_.IsEntra }; Descending = $true }, @{ e = { $_.LastUsed }; Descending = $true })
            if (-not $cands) { $cands = $profiles }
            $old = Select-FromList 'Whose profile is this laptop''s? (the person moving, not you)' $cands { param($x) '{0,-22} {1,-34} last used {2}' -f $x.Folder, $x.Account, $x.LastUsed }
            if (-not $old) { Write-Warn2 'Stopped. Nothing changed.'; return }
        }
        if (-not $old) { throw "No profile folder '$OldProfile'." }
        Write-Ok "Profile: $($old.Path) ($($old.Account))"

        $snap = Join-Path $WorkFolder "Prepare-$env:COMPUTERNAME-$script:Stamp.txt"
        $lines = [System.Collections.Generic.List[string]]::new()
        $lines.Add("Computer $env:COMPUTERNAME  Taken $(Get-Date)  Old tenant $($ds.TenantName) $($ds.TenantId)  DeviceId $($ds.DeviceId)")
        $lines.Add('===== BitLocker C: ====='); try { $lines.AddRange([string[]](manage-bde -protectors -get C: 2>&1)) } catch { $lines.Add("manage-bde: $_") }
        $lines.Add('===== Printers ====='); try { Get-Printer | ForEach-Object { $lines.Add("$($_.Name) | $($_.PortName) | $($_.DriverName)") } } catch {}
        $lines.Add('===== Mapped drives ====='); try { Get-SmbMapping -ErrorAction SilentlyContinue | ForEach-Object { $lines.Add("$($_.LocalPath) -> $($_.RemotePath)") } } catch {}
        $lines.Add('===== Profiles ====='); $profiles | ForEach-Object { $lines.Add("$($_.Folder) | $($_.Account) | $($_.Sid) | created $($_.Created)") }
        $lines | Set-Content $snap -Encoding UTF8 -WhatIf:$false
        if ($lines -match 'Numerical Password') { Write-Ok "BitLocker recovery key saved in $snap. Copy it to your password manager now." } else { Write-Warn2 'No BitLocker recovery password found (BitLocker off, or no recovery protector). Fine if BitLocker is off.' }

        Invoke-AutopilotCapture

        if (-not (Get-LocalUser -Name $LocalAdminName -ErrorAction SilentlyContinue)) {
            if ($PSCmdlet.ShouldProcess("Local admin '$LocalAdminName'", 'Create')) {
                $how = Select-FromList "Password for the local admin '$LocalAdminName'" @('Make an easy one for me, like Move4821! (shown once)', 'I will type my own') { param($x) $x }
                if (-not $how) { Write-Warn2 'Stopped. Nothing changed.'; return }
                if ($how -like 'Make*') {
                    $plain = New-EasyPassword
                    Write-Host "`n   $LocalAdminName password:  $plain`n" -ForegroundColor Black -BackgroundColor Yellow
                    Write-Host '   Write it down or save it now. It is not stored anywhere. Delete this account when the move is done.' -ForegroundColor Yellow
                    if (-not (Confirm-Go 'Saved it?' 'Yes, saved' -NoDefault)) { Write-Warn2 'Stopped. Nothing changed.'; return }
                    $pw = ConvertTo-SecureString $plain -AsPlainText -Force; $plain = $null
                } else { $pw = Read-Host "Type the password for '$LocalAdminName'" -AsSecureString }
                New-LocalUser -Name $LocalAdminName -Password $pw -PasswordNeverExpires -AccountNeverExpires -Description 'Tenant migration safety net' | Out-Null
                Add-LocalGroupMember -SID 'S-1-5-32-544' -Member $LocalAdminName
                Write-Ok "Local admin '$LocalAdminName' created"
            }
        } else { Write-Ok "Local admin '$LocalAdminName' already exists" }
        $whyNot = Test-LocalAdminReady $LocalAdminName
        if ($whyNot) {
            if ($WhatIfPreference) { Write-Warn2 "$whyNot A real run would stop here until it is fixed." }
            else {
                Set-State ([pscustomobject]@{ Stage = 'PreflightFailed'; Reason = $whyNot })
                throw "STOP: $whyNot Without a working local admin there is no way back in after leaving. Fix it (Enable-LocalUser / Add-LocalGroupMember) and run again."
            }
        } else { Write-Ok "Local admin '$LocalAdminName' is enabled and an administrator. Test it: sign in as it once before stage 2." }

        Set-State ([pscustomobject]@{ Stage = 'Prepared'; OldProfile = $old.Folder; OldSid = $old.Sid; OldPath = $old.Path; OldTenantId = $ds.TenantId; OldTenantName = $ds.TenantName; OldDeviceId = $ds.DeviceId; Snapshot = $snap; LocalAdmin = $LocalAdminName; Package = $PackagePath })
        $state = Get-State
    }

    if ($state.Stage -eq 'Prepared') {
        # ----- Stage 1b: leave
        if (-not $PackagePath -and -not $state.Package) {
            $found = @(Get-ChildItem $WorkFolder, $PSScriptRoot -Filter *.ppkg -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -Unique)
            $pick = Select-FromList 'Which enrolment package joins this laptop to the new tenant?' $found { param($x) '{0}   ({1})' -f $x.Name, $x.LastWriteTime }
            if ($pick) { Add-State 'Package' $pick.FullName }
            else { Write-Warn2 "No package chosen. Put the .ppkg in $WorkFolder before stage 2, or join by hand." }
        } elseif ($PackagePath) { Add-State 'Package' $PackagePath }
        Write-Host @"
About to leave $($state.OldTenantName). After this:
  - the person's old account can no longer sign in to this laptop (their files stay in $($state.OldPath))
  - sign in as .\$($state.LocalAdmin) and run this script again for stage 2
"@ -ForegroundColor Yellow
        if ($PSCmdlet.ShouldProcess("This laptop", "Leave tenant $($state.OldTenantName)")) {
            # Safety guard: re-check live, just before the point of no return
            $live = Get-DsregStatus
            if ($live.AzureAdJoined -ne 'YES') { throw 'This PC is no longer Microsoft Entra joined. Check dsregcmd /status before going further.' }
            if ($state.OldTenantId -and ($live.TenantId -ne $state.OldTenantId)) { throw "STOP: joined to $($live.TenantName) ($($live.TenantId)) now, but stage 1 recorded $($state.OldTenantId). Delete $script:StateFile and start again." }
            $whyNot = Test-LocalAdminReady $state.LocalAdmin
            if ($whyNot) { throw "STOP: $whyNot Not leaving the tenant without another way in." }
            if ($SourceTenantId) {
                if ($live.TenantId.Trim('{} ') -ne $SourceTenantId.Trim('{} ')) { throw "STOP: this machine is joined to $($live.TenantName) ($($live.TenantId)), not the source tenant $SourceTenantId. Nothing has changed." }
                if (-not (Confirm-Go "Take this laptop out of $($live.TenantName)?" "Yes, leave $($live.TenantName)" -NoDefault)) { Write-Warn2 'Stopped. Nothing changed.'; return }
            }
            else {
                if (-not $live.TenantName) { throw 'dsregcmd shows no tenant name to confirm. Run again with -SourceTenantId <source-guid>.' }
                Write-Host "This machine is joined to: $($live.TenantName)  ($($live.TenantId))" -ForegroundColor Yellow
                Write-Host 'If that is not the tenant you are migrating from, or this is your own PC, choose 0.' -ForegroundColor Yellow
                Write-Host "Computer: $env:COMPUTERNAME   Person: $($state.OldProfile)" -ForegroundColor Yellow
                if (-not (Confirm-Go "Take $env:COMPUTERNAME out of $($live.TenantName)?" "Yes, this is the person's laptop: leave $($live.TenantName)" -NoDefault)) { Write-Warn2 'Stopped. Nothing changed.'; return }
            }
            dsregcmd /leave | Out-Null
            Start-Sleep 5
            $after = Get-DsregStatus
            if ($after.AzureAdJoined -eq 'YES') { throw 'dsregcmd /leave did not take. Run this stage from a local admin session (not the Entra user) and try again.' }
            Add-State 'Stage' 'Left'
            Write-Ok 'Left the old tenant. Restart now, sign in as the local admin, run this script again.'
            if (Confirm-Go 'Restart now?' 'Yes, restart') { Restart-Computer -Force }
        }
        return
    }

    if ($state.Stage -eq 'Left') {
        # ----- Stage 2: join
        if ($ds.AzureAdJoined -eq 'YES') { Add-State 'Stage' 'Joined'; $state = Get-State }
        else {
            $pkg = if ($PackagePath) { $PackagePath } else { $state.Package }
            if (-not $pkg -or -not (Test-Path $pkg)) {
                # Not recorded at stage 1 (copied in later): look in the work folder and next to the script
                $found = @(Get-ChildItem $WorkFolder, $PSScriptRoot -Filter *.ppkg -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -Unique)
                $pick = Select-FromList 'Which enrolment package joins this laptop to the new tenant?' $found { param($x) '{0}   ({1})' -f $x.Name, $x.LastWriteTime }
                $pkg = if ($pick) { $pick.FullName } else { $null }
                if (-not $found) { Write-Warn2 "No .ppkg in $WorkFolder. Copy it there and run again, or join by hand (below)." }
            }
            if ($pkg -and (Test-Path $pkg)) {
                if ($PSCmdlet.ShouldProcess("This laptop", "Join the new tenant with $pkg")) {
                    Import-Module Provisioning -ErrorAction Stop
                    Install-ProvisioningPackage -PackagePath $pkg -QuietInstall -ForceInstall | Out-Null
                    Start-Sleep 20
                    $after = Get-DsregStatus
                    if ($after.AzureAdJoined -eq 'YES') { Write-Ok "Joined $($after.TenantName) ($($after.TenantId))"; Add-State 'Stage' 'Joined'; Add-State 'NewTenantId' $after.TenantId }
                    else { Write-Warn2 'Not joined yet. Provisioning can take a minute. Restart, sign in as the local admin and run again; if still not joined, join by hand (below).' }
                }
            }
            if ((Get-State).Stage -ne 'Joined') {
                Write-Host @"
Join by hand: Settings > Accounts > Access work or school > Connect > "Join this device to Microsoft Entra ID",
sign in with the person's NEW work account, restart, sign in as .\$($state.LocalAdmin), run this script again.
"@ -ForegroundColor Yellow
                return
            }
        }
    }

    if ((Get-State).Stage -eq 'Joined') {
        Write-Step 'Stage 3: first sign-in'
        Write-Host @"
1. Sign out. The person signs in with their NEW email address and password ("Other user" on the sign-in screen).
   This takes a few minutes and creates their new account on the laptop. Let them set a PIN if asked.
2. They sign out. You sign in as .\$($state.LocalAdmin) and run this script again, option 5 (hand the old profile over).
"@ -ForegroundColor Yellow
        Add-State 'Stage' 'AwaitingFirstSignIn'
        return
    }
    if ((Get-State).Stage -in 'AwaitingFirstSignIn', 'HandoverFailed') { Invoke-HandoverProfile; return }
    if ((Get-State).Stage -eq 'Done') { Write-Ok "This laptop is done. Report: $($state.Report)" }
}

# ---------------------------------------------------------------- 5. profile handover
function Clear-OldSignIns([string]$ProfilePath, [string]$Log) {
    # The handed-over profile still caches sign-ins for the OLD tenant: the Windows sign-in broker (WAM), Office's
    # identity list and Teams. Same email, different tenant, so apps loop on "Something went wrong" or use stale
    # tokens. Everything is renamed or exported first, so it can be put back. Files and OneDrive data are not touched.
    $stamp = Get-Date -Format yyyyMMddHHmm
    $done = @()
    $folders = @(
        'AppData\Local\Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\AC\TokenBroker\Accounts',
        'AppData\Local\Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\AC\TokenBroker\Cache',
        'AppData\Local\Microsoft\TokenBroker\Cache',
        'AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache',
        'AppData\Roaming\Microsoft\Teams',
        'AppData\Local\Microsoft\IdentityCache',
        'AppData\Local\Microsoft\OneAuth'
    )
    foreach ($f in $folders) {
        $full = Join-Path $ProfilePath $f
        if (Test-Path $full) {
            try { Rename-Item $full "$(Split-Path $full -Leaf).old-$stamp" -ErrorAction Stop; $done += $f } catch { Write-Warn2 "Could not rename $f : $($_.Exception.Message)" }
        }
    }
    # Office identities and cached sign-in state in the person's registry hive
    $hive = Join-Path $ProfilePath 'NTUSER.DAT'
    $key = "MIG_$([guid]::NewGuid().ToString('N').Substring(0,8))"
    reg load "HKU\$key" "$hive" | Out-Null
    $eap = $ErrorActionPreference; $ErrorActionPreference = 'Continue'   # reg.exe writes to stderr for missing keys
    try {
        foreach ($k in 'Software\Microsoft\Office\16.0\Common\Identity', 'Software\Microsoft\Office\16.0\Common\Internet\WebServiceCache', 'Software\Microsoft\Office\16.0\Common\ServicesManagerCache') {
            $rk = "HKU\$key\$k"
            if (Test-Path "Registry::HKEY_USERS\$key\$k") {
                $bak = Join-Path $WorkFolder ("RegBackup-$stamp-" + ($k -replace '[\\ ]', '_') + '.reg')
                reg export $rk $bak /y | Out-Null
                reg delete $rk /f | Out-Null
                $done += "HKCU\$k (backup $bak)"
            }
        }
    }
    finally { $ErrorActionPreference = $eap; [gc]::Collect(); Start-Sleep 1; reg unload "HKU\$key" 2>&1 | Out-Null }
    if ($done) { "Cleared old sign-ins:`r`n  " + ($done -join "`r`n  ") | Add-Content $Log -WhatIf:$false; Write-Ok "Cleared $($done.Count) cached old-tenant sign-in item(s). Apps will ask for the new sign-in once." }
    else { Write-Ok 'No cached old-tenant sign-ins found.' }
}
function Grant-HiveAccess([string]$HivePath, [string]$Sid) {
    # Loads a registry hive file and grants the SID full control at the root with inheritance
    $key = "MIG_$([guid]::NewGuid().ToString('N').Substring(0,8))"
    reg load "HKU\$key" "$HivePath" | Out-Null
    try {
        $rk = [Microsoft.Win32.Registry]::Users.OpenSubKey($key, [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree, [System.Security.AccessControl.RegistryRights]::ChangePermissions -bor [System.Security.AccessControl.RegistryRights]::TakeOwnership -bor [System.Security.AccessControl.RegistryRights]::ReadKey)
        $acl = $rk.GetAccessControl()
        $rule = New-Object System.Security.AccessControl.RegistryAccessRule(([System.Security.Principal.SecurityIdentifier]$Sid), 'FullControl', 'ContainerInherit', 'None', 'Allow')
        $acl.AddAccessRule($rule); $rk.SetAccessControl($acl); $rk.Close()
    }
    finally { [gc]::Collect(); Start-Sleep 1; reg unload "HKU\$key" | Out-Null }
}
function Invoke-HandoverProfile {
    if (-not (Test-Elevated)) { throw 'Run as administrator.' }
    Write-Step 'Hand the old profile to the new account'
    $state = Get-State
    $profiles = @(Get-Profiles)
    if (-not $OldProfile) { $OldProfile = $state.OldProfile }
    if ($OldProfile) { $old = $profiles | Where-Object { $_.Folder -eq (Split-Path $OldProfile -Leaf) } | Select-Object -First 1 }
    else { $old = Select-FromList 'Which is the OLD profile (the one with their files)?' @($profiles | Where-Object { $_.Folder -ne $LocalAdminName }) { param($x) '{0,-22} {1,-34} created {2}' -f $x.Folder, $x.Account, $x.Created } }
    if (-not $old) { throw 'No old profile chosen.' }
    if ($NewProfile) { $new = $profiles | Where-Object { $_.Folder -eq (Split-Path $NewProfile -Leaf) } | Select-Object -First 1 }
    else {
        $cands = @($profiles | Where-Object { $_.IsEntra -and $_.Sid -ne $old.Sid -and $_.Created -gt $old.Created -and $_.Folder -ne $LocalAdminName } | Sort-Object Created -Descending)
        if (-not $cands) { throw 'No new Microsoft Entra profile found. The person must sign in once with the new account first, then sign out.' }
        $new = Select-FromList "Which is the person's NEW account (their new email)? Not yours." $cands { param($x) '{0,-22} {1,-34} created {2}' -f $x.Folder, $x.Account, $x.Created }
    }
    if (-not $new) { Write-Warn2 'Stopped. Nothing changed.'; return }
    if ($new.Loaded -or $old.Loaded) {
        $busy = @($old, $new | Where-Object Loaded | ForEach-Object { $_.Folder }) -join ', '
        Write-Warn2 "Windows still has $busy open (someone is signed in, used Switch user, or Windows has not let go of it yet)."
        Write-Host 'Fix: restart, sign in as .\' -NoNewline -ForegroundColor Yellow; Write-Host "$LocalAdminName ONLY (the person must not sign in first), run this again." -ForegroundColor Yellow
        if (Confirm-Go 'Restart now?' 'Yes, restart now') { Restart-Computer -Force }
        return
    }
    Write-Host "Old: $($old.Path)  $($old.Account)  $($old.Sid)" -ForegroundColor Yellow
    Write-Host "New: $($new.Path)  $($new.Account)  $($new.Sid)" -ForegroundColor Green
    if (-not $PSCmdlet.ShouldProcess("$($new.Account)", "Point at $($old.Path) and take over its permissions")) { return }
    if (-not (Confirm-Go "Give $($new.Account) the files and settings in $($old.Path)?" 'Yes, hand it over' -NoDefault)) { Write-Warn2 'Stopped. Nothing changed.'; return }

    $plKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$($new.Sid)"
    $parked = "$($new.Path).new"
    $log = Join-Path $WorkFolder "Handover-$env:COMPUTERNAME-$script:Stamp.log"
    try {
        # 1. Park the empty new profile folder, point the new SID at the old folder
        Rename-Item $new.Path $parked -ErrorAction Stop
        Set-ItemProperty $plKey -Name ProfileImagePath -Value $old.Path -ErrorAction Stop
        "ProfileImagePath for $($new.Sid) -> $($old.Path); new folder parked as $parked" | Add-Content $log
        # 2. File permissions: new SID becomes owner with full control, old SID stays (harmless)
        $sidRef = "*$($new.Sid)"
        # icacls reports "Access is denied" for the legacy junctions every profile has (Application Data, Cookies...).
        # Expected and harmless. Windows PowerShell turns native stderr into terminating errors under 'Stop', so relax it here.
        $count = (Get-ChildItem $old.Path -Recurse -Force -File -ErrorAction SilentlyContinue | Measure-Object).Count
        Write-Host "Changing owner and permissions on $count files in two passes. This shows nothing while it works and can take 10 to 30 minutes on a full profile. Do not close this window." -ForegroundColor Yellow
        $eap = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
        try {
            icacls "$($old.Path)" /setowner "$sidRef" /T /C /Q 2>&1 | ForEach-Object { "$_" } | Add-Content $log -WhatIf:$false
            Write-Host "   Pass 1 of 2 done (owner) $(Get-Date -Format HH:mm). Now pass 2 (permissions)..." -ForegroundColor Cyan
            icacls "$($old.Path)" /grant "${sidRef}:(OI)(CI)F" /T /C /Q 2>&1 | ForEach-Object { "$_" } | Add-Content $log -WhatIf:$false
        }
        finally { $ErrorActionPreference = $eap }
        # The folder itself must now belong to the new account; anything inside that was skipped is logged
        $ownerSid = (Get-Acl $old.Path).GetOwner([Security.Principal.SecurityIdentifier]).Value
        if ($ownerSid -ne $new.Sid) { throw "Ownership did not change on $($old.Path) (owner is $ownerSid). See $log" }
        # 3. Registry hives inside the profile
        Grant-HiveAccess (Join-Path $old.Path 'NTUSER.DAT') $new.Sid
        $usr = Join-Path $old.Path 'AppData\Local\Microsoft\Windows\UsrClass.dat'
        if (Test-Path $usr) { Grant-HiveAccess $usr $new.Sid }
        # 4. Verify
        $check = Get-CimInstance Win32_UserProfile | Where-Object SID -eq $new.Sid
        if ($check.LocalPath -ne $old.Path) { throw "Verification failed: profile for $($new.Sid) is $($check.LocalPath)" }
        Write-Ok "Done. $($new.Account) now uses $($old.Path). Empty profile parked at $parked (delete it after 30 days)."
        # 5. Clear the old tenant's cached sign-ins (renamed / exported, reversible)
        try { Clear-OldSignIns $old.Path $log } catch { Write-Warn2 "Clearing old sign-ins: $($_.Exception.Message). Not critical: sign out of Office, Teams and OneDrive once and sign in again." }
        Write-Host @'
Next, signed in as the person:
  - Office (Word/Outlook) and Teams: sign in once with the new email.
  - OneDrive: click the cloud icon > Settings > Account > Unlink this PC, then sign in with the new email.
    The old OneDrive folder stays on disk as a local copy; their files are already in the new OneDrive.
  - Outlook classic: if it keeps showing errors, Control Panel > Mail > Show Profiles > Add a new profile.
'@ -ForegroundColor Cyan
        Write-Host 'Saved browser passwords and Windows Hello are tied to the old account and need setting up again.' -ForegroundColor Yellow
        $rep = Join-Path $WorkFolder "DeviceMove-$env:COMPUTERNAME-$script:Stamp.csv"
        [pscustomobject]@{ Computer = $env:COMPUTERNAME; Status = 'Done'; OldTenant = $state.OldTenantName; NewTenant = (Get-DsregStatus).TenantName; Person = $new.Account; ProfilePath = $old.Path; Parked = $parked; Method = 'Handover'; When = Get-Date } | Export-Csv $rep -NoTypeInformation -WhatIf:$false
        Add-State 'Stage' 'Done'; Add-State 'Report' $rep
        Invoke-BitLockerBackup
    }
    catch {
        Write-Bad "Handover failed: $($_.Exception.Message). Rolling back."
        try { Set-ItemProperty $plKey -Name ProfileImagePath -Value $new.Path -ErrorAction SilentlyContinue; if (Test-Path $parked) { Rename-Item $parked $new.Path -ErrorAction SilentlyContinue } } catch {}
        Add-State 'Stage' 'HandoverFailed'
        Write-Host 'The old profile is untouched. Use option 6 (copy the data into the new profile) instead.' -ForegroundColor Yellow
    }
}
function Invoke-BitLockerBackup {
    try {
        $vol = Get-BitLockerVolume -MountPoint 'C:' -ErrorAction Stop
        $rp = $vol.KeyProtector | Where-Object KeyProtectorType -eq 'RecoveryPassword' | Select-Object -First 1
        if ($rp -and (Get-DsregStatus).AzureAdJoined -eq 'YES') { BackupToAAD-BitLockerKeyProtector -MountPoint 'C:' -KeyProtectorId $rp.KeyProtectorId -ErrorAction Stop | Out-Null; Write-Ok 'BitLocker recovery key backed up to the new tenant' }
    } catch { Write-Warn2 "BitLocker key not backed up to the new tenant: $($_.Exception.Message)" }
}

# ---------------------------------------------------------------- 6. copy fallback
function Invoke-CopyProfile {
    $s = Get-SiblingScript 'Copy-ProfileData.ps1'
    $state = Get-State
    $op = if ($OldProfile) { $OldProfile } else { $state.OldProfile }
    $p = @{ BackupFolder = $WorkFolder }
    if ($op) { $p.OldProfile = $op }
    if ($NewProfile) { $p.NewProfile = $NewProfile }
    if ($WhatIfPreference) { $p.WhatIf = $true }
    & $s @p
    if (-not $WhatIfPreference) { Add-State 'Stage' 'Done'; Invoke-BitLockerBackup }
}

# ---------------------------------------------------------------- update
function Invoke-SelfUpdate {
    # Short to type on any machine: .\Move-TenantDevices.ps1 -Action Update
    $here = Split-Path $PSCommandPath
    foreach ($n in 'Move-TenantDevices.ps1', 'Copy-ProfileData.ps1') {
        $tmp = Join-Path $here "$n.new"
        Invoke-WebRequest "$script:Repo/$n`?nocache=$(Get-Random)" -OutFile $tmp -UseBasicParsing -Headers @{ 'Cache-Control' = 'no-cache' }
        Move-Item $tmp (Join-Path $here $n) -Force -WhatIf:$false
        Unblock-File (Join-Path $here $n) -ErrorAction SilentlyContinue
    }
    $v = (Select-String 'Version : ' (Join-Path $here 'Move-TenantDevices.ps1') | Select-Object -First 1).Line.Trim()
    Write-Ok "Updated. $v"
}

# ---------------------------------------------------------------- 7. clean up the admin PC
function Invoke-CleanupAdminPC {
    Write-Step 'Clean up this admin PC (what option 1 installed or left behind)'
    $items = @()
    $tk = Join-Path $env:LOCALAPPDATA 'MigrationToolkit'
    if (Test-Path $tk) { $items += [pscustomobject]@{ What = "Private designer copies and temp files ($tk)"; Do = { Remove-Item $tk -Recurse -Force } } }
    $fe = 'HKCU:\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION'
    if ((Get-ItemProperty $fe -Name 'powershell.exe' -ErrorAction SilentlyContinue)) { $items += [pscustomobject]@{ What = 'Sign-in window rendering setting for powershell.exe (HKCU)'; Do = { Remove-ItemProperty $fe -Name 'powershell.exe' } } }
    $logs = @(Get-ChildItem $WorkFolder -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -like 'ICD-build-*.log*' -or $_.Name -like '*.customizations.xml' })
    if ($logs) { $items += [pscustomobject]@{ What = "$($logs.Count) build log(s) in $WorkFolder (your .ppkg files are kept)"; Do = { $logs | Remove-Item -Force } } }
    $pkg = Get-AppxPackage -Name '*WindowsConfigurationDesigner*' -ErrorAction SilentlyContinue
    if ($pkg) { $items += [pscustomobject]@{ What = "Windows Configuration Designer (Store app $($pkg[0].Version))"; Do = {
        try { $pkg | Remove-AppxPackage } catch { winget uninstall --id 9NBLGGH4TX22 --source msstore --silent --accept-source-agreements | Out-Host } } } }
    if (Find-Wcd | Where-Object Source -eq 'ADK') { $items += [pscustomobject]@{ What = 'Windows ADK (Configuration Designer). Windows asks once to allow the uninstall'; Do = {
        winget uninstall --id Microsoft.WindowsADK --silent --accept-source-agreements | Out-Host } } }

    if (-not $items) { Write-Ok 'Nothing to clean up.'; return }
    Write-Host 'This will remove:' -ForegroundColor Cyan
    $items | ForEach-Object { Write-Host "  - $($_.What)" }
    Write-Host 'It does not touch: your .ppkg files, Microsoft Graph or other PowerShell modules, or anything in either tenant.' -ForegroundColor Gray
    if ($WhatIfPreference) { Write-Host 'WhatIf: nothing removed.' -ForegroundColor Yellow; return }
    if (-not (Confirm-Go 'Remove all of the above?' 'Yes, remove them')) { return }
    foreach ($i in $items) {
        try { & $i.Do; Write-Ok "Removed: $($i.What)" } catch { Write-Warn2 "$($i.What): $($_.Exception.Message)" }
    }
    $left = @(Get-ChildItem $WorkFolder -Filter *.ppkg -ErrorAction SilentlyContinue)
    if ($left) { Write-Host "Kept $($left.Count) package(s) in $WorkFolder. Delete them yourself once every laptop has moved: they join devices to the tenant until the token expires." -ForegroundColor Yellow }
}

# ---------------------------------------------------------------- menu
if ($Action -eq 'Menu') {
    $st = Get-State
    Write-Host @"

 Move-TenantDevices  (state on this machine: $($st.Stage))
 -------------------------------------------------------------
  Any admin PC (your own laptop is fine)
   1  Build the enrolment package for the destination tenant
   2  Copy Intune configuration from source to destination
   3  Autopilot: capture this device's hash / move hashes
  Laptop being moved only (run as administrator; never option 4 on your own PC)
   4  Move this laptop to the new tenant (staged, re-run after each restart)
   5  Hand the old profile to the new account (after first sign-in)
   6  Copy profile data into the new profile instead (fallback)
  Admin PC, when you are finished
   7  Clean up this admin PC (removes what option 1 installed)
   8  Update this script to the latest version
   0  Exit
"@ -ForegroundColor Cyan
    switch (Read-Choice 'Choose' @('0', '1', '2', '3', '4', '5', '6', '7', '8')) {
        '1' { $Action = 'BuildPackage' }
        '2' { $Action = 'CopyIntune' }
        '3' { $Action = if ((Read-Choice 'Capture this device (C) or move captured hashes between tenants (M)?' @('C', 'M', 'c', 'm')) -in 'C', 'c') { 'AutopilotCapture' } else { 'AutopilotMove' } }
        '4' { $Action = 'MoveDevice' }
        '5' { $Action = 'HandoverProfile' }
        '6' { $Action = 'CopyProfile' }
        '7' { $Action = 'CleanupAdminPC' }
        '8' { $Action = 'Update' }
        default { return }
    }
}
switch ($Action) {
    'BuildPackage'     { Invoke-BuildPackage }
    'CopyIntune'       { Invoke-CopyIntune }
    'AutopilotCapture' { if (-not (Test-Elevated)) { throw 'Run as administrator.' }; Invoke-AutopilotCapture }
    'AutopilotMove'    { Invoke-AutopilotMove }
    'MoveDevice'       { Invoke-MoveDevice }
    'HandoverProfile'  { Invoke-HandoverProfile }
    'CopyProfile'      { Invoke-CopyProfile }
    'CleanupAdminPC'   { Invoke-CleanupAdminPC }
    'Update'           { Invoke-SelfUpdate }
}
