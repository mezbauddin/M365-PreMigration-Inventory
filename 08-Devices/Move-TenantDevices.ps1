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
                                                      valid up to 180 days. Falls back to Windows Configuration Designer.
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

.PARAMETER Action
    Skip the menu: BuildPackage, CopyIntune, AutopilotCapture, AutopilotMove, MoveDevice, HandoverProfile, CopyProfile.

.PARAMETER SourceTenantId
    Source tenant ID (GUID). CopyIntune and AutopilotMove sign in to it. MoveDevice uses it as a safety check: if
    this machine's dsregcmd TenantId is different, it stops before changing anything. Strongly recommended for
    MoveDevice.

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
    Version : 1.0.5
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: Windows PowerShell 5.1 or 7, elevated. Admin-PC options need Microsoft.Graph.Authentication.
              Laptop options need nothing extra (Provisioning module is built in).
    Status  : Option 1's scripted token request uses the same sign-in flow Windows Configuration Designer uses, which
              Microsoft doesn't document. If it stops working, the script opens WCD for you instead. Options 2, 3 and 5
              are tested against stubbed tenants and devices; run them with -WhatIf first on anything that matters.
    Changelog:
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
    [ValidateSet('Menu', 'BuildPackage', 'CopyIntune', 'AutopilotCapture', 'AutopilotMove', 'MoveDevice', 'HandoverProfile', 'CopyProfile')]
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
    [switch]$UseDeviceCode,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$script:Repo = 'https://raw.githubusercontent.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit/main/08-Devices'
if (-not $WorkFolder) { $WorkFolder = if (Test-Path 'C:\') { 'C:\MigrationBackup' } else { Join-Path (Get-Location) 'DeviceMigration' } }
New-Item -ItemType Directory -Path $WorkFolder -Force -WhatIf:$false | Out-Null
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
        [pscustomobject]@{ Folder = Split-Path $_.LocalPath -Leaf; Path = $_.LocalPath; Account = $name; Sid = $sid
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
    # Modern, interactive sign-in for the device registration service, so MFA and Conditional Access apply and
    # tenants that block device code still work. Uses MSAL from the Microsoft.Graph.Authentication module.
    # Tries, in order: the Windows sign-in broker (WAM), the system browser, then the embedded window
    # (Windows PowerShell only). Each is a Microsoft first-party public client; which one a tenant accepts depends on
    # its policies, so the script reports each attempt.
    $drsScope = 'urn:ms-drs:enterpriseregistration.windows.net/.default'
    if (-not ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'Microsoft.Identity.Client' -and $_.GetType('Microsoft.Identity.Client.PublicClientApplicationBuilder') })) {
        $mod = Get-Module -ListAvailable Microsoft.Graph.Authentication | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $mod) { throw 'Install-Module Microsoft.Graph.Authentication -Scope CurrentUser first (it provides the Microsoft sign-in library).' }
        $edition = if ($PSVersionTable.PSEdition -eq 'Core') { 'Core' } else { 'Desktop' }
        $dlls = @(Get-ChildItem $mod.ModuleBase -Recurse -Filter 'Microsoft.Identity.Client*.dll' -ErrorAction SilentlyContinue)
        $pick = { param($n) $c = @($dlls | Where-Object Name -eq $n); $e = @($c | Where-Object { $_.FullName -match "\\$edition\\" }); if ($e) { $e[0] } elseif ($c) { $c[0] } }
        foreach ($n in 'Microsoft.IdentityModel.Abstractions.dll', 'Microsoft.Identity.Client.dll', 'Microsoft.Identity.Client.NativeInterop.dll', 'Microsoft.Identity.Client.Broker.dll', 'Microsoft.Identity.Client.Desktop.dll') {
            $f = & $pick $n
            if (-not $f -and $n -eq 'Microsoft.IdentityModel.Abstractions.dll') { $f = Get-ChildItem $mod.ModuleBase -Recurse -Filter $n -ErrorAction SilentlyContinue | Select-Object -First 1 }
            if ($f) { try { Add-Type -Path $f.FullName -ErrorAction Stop } catch {} }
        }
    }
    if (-not ('Win32Mig.Native' -as [type])) {
        Add-Type -Namespace Win32Mig -Name Native -MemberDefinition '[DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow(); [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();'
    }
    $hwnd = [Win32Mig.Native]::GetForegroundWindow(); if ($hwnd -eq [IntPtr]::Zero) { $hwnd = [Win32Mig.Native]::GetConsoleWindow() }
    $authority = "https://login.microsoftonline.com/$TenantId"
    $wcd = 'b90d5b8f-5503-4153-b545-b31cecfaece2'      # Windows Configuration Designer
    $aadps = '1b730954-1685-4b74-9bfd-dac224a7b894'    # Azure AD PowerShell (legacy first-party client used for device registration tokens)
    $attempts = @(
        @{ Label = 'Windows sign-in broker (WAM)'; Client = $wcd; Mode = 'Broker' },
        @{ Label = 'system browser'; Client = $wcd; Mode = 'Browser' },
        @{ Label = 'system browser (device registration client)'; Client = $aadps; Mode = 'Browser' },
        @{ Label = 'sign-in window'; Client = $aadps; Mode = 'Embedded' }
    )
    foreach ($a in $attempts) {
        if ($a.Mode -eq 'Embedded' -and $PSVersionTable.PSEdition -eq 'Core') { continue }
        try {
            Write-Host "Signing in with the $($a.Label)..." -ForegroundColor Cyan
            $b = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($a.Client).WithAuthority($authority)
            switch ($a.Mode) {
                'Broker' {
                    $opt = [Microsoft.Identity.Client.BrokerOptions]::new([Microsoft.Identity.Client.BrokerOptions+OperatingSystems]::Windows)
                    $b = [Microsoft.Identity.Client.Broker.BrokerExtension]::WithBroker($b, $opt)
                }
                'Browser'  { $b = $b.WithRedirectUri('http://localhost') }
                'Embedded' { $b = $b.WithRedirectUri('urn:ietf:wg:oauth:2.0:oob') }
            }
            $app = $b.Build()
            $req = $app.AcquireTokenInteractive([string[]]@($drsScope)).WithPrompt([Microsoft.Identity.Client.Prompt]::SelectAccount).WithParentActivityOrWindow($hwnd)
            if ($Admin) { $req = $req.WithLoginHint($Admin) }
            if ($a.Mode -eq 'Embedded') { $req = $req.WithUseEmbeddedWebView($true) }
            $res = $req.ExecuteAsync().GetAwaiter().GetResult()
            if ($res.AccessToken) { Write-Ok "Signed in as $($res.Account.Username) ($($a.Label))"; return $res.AccessToken }
        }
        catch { Write-Warn2 "$($a.Label): $(($_.Exception.InnerException, $_.Exception | Where-Object { $_ } | Select-Object -First 1).Message)" }
    }
    throw 'No interactive sign-in method was accepted by this tenant.'
}
function Get-DrsAccessTokenDeviceCode([string]$TenantId, [string]$Admin) {
    # Only with -UseDeviceCode: many tenants block device code flow by Conditional Access
    $client = 'b90d5b8f-5503-4153-b545-b31cecfaece2'; $drs = 'urn:ms-drs:enterpriseregistration.windows.net'
    $dc = Invoke-RestMethod -Method Post "https://login.microsoftonline.com/$TenantId/oauth2/devicecode" -Body @{ client_id = $client; resource = $drs }
    Write-Host $dc.message -ForegroundColor Yellow
    $deadline = (Get-Date).AddSeconds([int]$dc.expires_in)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds ([int]$dc.interval)
        try { $t = Invoke-RestMethod -Method Post "https://login.microsoftonline.com/$TenantId/oauth2/token" -Body @{ grant_type = 'device_code'; client_id = $client; code = $dc.device_code; resource = $drs }; if ($t.access_token) { return $t.access_token } }
        catch { $bd = ''; try { $bd = $_.ErrorDetails.Message } catch {}; if ($bd -notmatch 'authorization_pending') { throw } }
    }
    throw 'Sign-in timed out.'
}
function Get-BulkToken([string]$TenantId, [string]$Admin, [int]$Days, [string]$PackageName) {
    # The same call Windows Configuration Designer makes behind "Get Bulk Token": it creates the package_<guid>
    # account in the tenant and returns a bulk token that lets devices join without a user signing in.
    $at = if ($UseDeviceCode) { Get-DrsAccessTokenDeviceCode $TenantId $Admin } else { Get-DrsAccessToken $TenantId $Admin }
    $h = @{ Authorization = "Bearer $at" }
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
function Invoke-Icd($Wcd, [string[]]$IcdArgs, [string]$Log) {
    # ADK version: run ICD.exe directly. Store version: run it inside its own package context.
    if ($Wcd.Source -eq 'ADK') { & $Wcd.Exe @IcdArgs 2>&1 | Out-String | Add-Content $Log; return }
    $pkg = Get-AppxPackage -Name '*WindowsConfigurationDesigner*' | Sort-Object Version -Descending | Select-Object -First 1
    $appId = try { @((Get-AppxPackageManifest $pkg).Package.Applications.Application)[0].Id } catch { 'App' }
    Invoke-CommandInDesktopPackage -PackageFamilyName $pkg.PackageFamilyName -AppId $appId -Command $Wcd.Exe -Args ($IcdArgs -join ' ') -PreventBreakaway
    # Invoke-CommandInDesktopPackage doesn't wait: poll for the package for up to 3 minutes
    $out = ($IcdArgs | Where-Object { $_ -like '/PackagePath:*' }) -replace '^/PackagePath:"?|"$', ''
    $until = (Get-Date).AddMinutes(3)
    while (-not (Test-Path $out) -and (Get-Date) -lt $until) { Start-Sleep -Seconds 5 }
}
function Install-AdkIcd {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) { return $null }
    Write-Host 'Installing the Windows ADK "Configuration Designer" feature only (silent, a few minutes)...' -ForegroundColor Cyan
    winget install --id Microsoft.WindowsADK --accept-package-agreements --accept-source-agreements --silent --override '/features OptionId.ICDConfigurationDesigner /quiet /norestart /ceip off' | Out-Host
    $w = Find-Wcd
    if ($w -and $w.Source -eq 'ADK') { Write-Ok 'Configuration Designer (ADK) installed'; return $w }
    $null
}

function Invoke-BuildPackage {
    Write-Step 'Build the enrolment package (run on your admin PC)'
    $wcd = Find-Wcd
    if ($wcd) { Write-Ok "Windows Configuration Designer found ($($wcd.Source) version)" } else { $wcd = Install-Wcd; if (-not $wcd) { return } }

    if (-not $DestinationTenantId) { $DestinationTenantId = Read-Host 'Destination tenant ID (GUID)' }
    if (-not $DestinationAdmin) { $DestinationAdmin = Read-Host 'Destination admin sign-in (e.g. admin@contoso.com)' }
    $name = Read-Host 'Package name (letters, digits and hyphens, e.g. contoso-join)'
    if (-not $name) { $name = 'tenant-join' }
    $name = $name -replace '[^A-Za-z0-9-]', ''
    $expiry = Read-Host 'Token validity in days (max 180, default 180)'
    if (-not $expiry) { $expiry = 180 }
    $expiry = [Math]::Min(180, [int]$expiry)
    $rename = Read-Host 'Rename laptops on join? Pattern like TG-%SERIAL% (max 15 characters in total), or Enter to keep names'
    $ppkg = Join-Path $WorkFolder "$name.ppkg"
    $log = Join-Path $WorkFolder "ICD-build-$script:Stamp.log"

    $built = $false
    $bprt = $null
    try { $bprt = Get-BulkToken $DestinationTenantId $DestinationAdmin $expiry $name; Write-Ok "Bulk token received (valid $expiry days)" }
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
        $attempts = @($wcd)
        foreach ($w in $attempts + @('ADK')) {
            if ($w -eq 'ADK') { if ($wcd.Source -eq 'ADK') { break }; $w = Install-AdkIcd; if (-not $w) { break } }
            if (-not $w.Exe) { continue }
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

    # Guided: Windows Configuration Designer's own wizard (requests its own token)
    Write-Step 'Guided: build it in Windows Configuration Designer (about 3 minutes)'
    Write-Host @"
1. Click "Provision desktop devices".
2. Name: $name. Next.
3. Set up device: leave the device name BLANK (a pattern renames every laptop and adds a restart). Turn off "Remove pre-installed software". Next.
4. Set up network: turn it OFF. Next.
5. Account management: choose "Enroll in Azure AD". Bulk Token Expiry: $expiry days from today. Click "Get Bulk Token" and sign in as $DestinationAdmin.
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
        if ($env:USERNAME -and (Get-Profiles | Where-Object { $_.Loaded -and $_.Account -like 'AzureAD\*' })) { Write-Warn2 'A Microsoft Entra user is signed in. Finish this stage, then restart and sign in as the local admin for stage 2.' }

        $profiles = @(Get-Profiles)
        $profiles | Format-Table Folder, Account, Created, LastUsed -AutoSize | Out-String -Width 200 | Write-Host
        if (-not $OldProfile) { $OldProfile = Read-Host 'Which profile folder belongs to the person who uses this laptop? (Folder column)' }
        $old = $profiles | Where-Object { $_.Folder -eq (Split-Path $OldProfile -Leaf) } | Select-Object -First 1
        if (-not $old) { throw "No profile folder '$OldProfile'." }

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
                $pw = Read-Host "Password for the local admin '$LocalAdminName' (keep it in your password manager)" -AsSecureString
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
            $p = Read-Host 'Path to the enrolment package (.ppkg) from option 1 (Enter to join by hand later)'
            if ($p) { if (-not (Test-Path $p)) { throw "$p not found" }; Add-State 'Package' $p }
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
                if ((Read-Choice 'Type LEAVE to continue, or N to stop' @('LEAVE', 'N', 'n')) -ne 'LEAVE') { return }
            }
            else {
                if (-not $live.TenantName) { throw 'dsregcmd shows no tenant name to confirm. Run again with -SourceTenantId <source-guid>.' }
                Write-Host "This machine is joined to: $($live.TenantName)  ($($live.TenantId))" -ForegroundColor Yellow
                Write-Host 'If that is not the source tenant you are migrating from, or this is your own PC, type N.' -ForegroundColor Yellow
                $typed = Read-Host "Type the tenant name exactly ($($live.TenantName)) to leave it, or N to stop"
                if ($typed.Trim() -cne $live.TenantName) { Write-Warn2 'Name did not match. Nothing changed.'; return }
            }
            dsregcmd /leave | Out-Null
            Start-Sleep 5
            $after = Get-DsregStatus
            if ($after.AzureAdJoined -eq 'YES') { throw 'dsregcmd /leave did not take. Run this stage from a local admin session (not the Entra user) and try again.' }
            Add-State 'Stage' 'Left'
            Write-Ok 'Left the old tenant. Restart now, sign in as the local admin, run this script again.'
            if ((Read-Choice 'Restart now? (Y/N)' @('Y', 'N', 'y', 'n')) -in 'Y', 'y') { Restart-Computer -Force }
        }
        return
    }

    if ($state.Stage -eq 'Left') {
        # ----- Stage 2: join
        if ($ds.AzureAdJoined -eq 'YES') { Add-State 'Stage' 'Joined'; $state = Get-State }
        else {
            $pkg = if ($PackagePath) { $PackagePath } else { $state.Package }
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
    if (-not $OldProfile) { $profiles | Format-Table Folder, Account, Created -AutoSize | Out-String -Width 200 | Write-Host; $OldProfile = Read-Host 'OLD profile folder name' }
    $old = $profiles | Where-Object { $_.Folder -eq (Split-Path $OldProfile -Leaf) } | Select-Object -First 1
    if (-not $old) { throw "Old profile '$OldProfile' not found." }
    $new = if ($NewProfile) { $profiles | Where-Object { $_.Folder -eq (Split-Path $NewProfile -Leaf) } | Select-Object -First 1 }
           else { $profiles | Where-Object { $_.Created -gt $old.Created -and $_.Account -like 'AzureAD\*' -and $_.Sid -ne $old.Sid } | Sort-Object Created | Select-Object -Last 1 }
    if (-not $new) { throw 'No new Microsoft Entra profile found. The person must sign in once with the new account first.' }
    if ($new.Loaded -or $old.Loaded) { throw 'A profile is still loaded (someone is signed in). Sign them out and run from the local admin.' }
    Write-Host "Old: $($old.Path)  $($old.Account)  $($old.Sid)" -ForegroundColor Yellow
    Write-Host "New: $($new.Path)  $($new.Account)  $($new.Sid)" -ForegroundColor Green
    if (-not $PSCmdlet.ShouldProcess("$($new.Account)", "Point at $($old.Path) and take over its permissions")) { return }
    if ((Read-Choice 'Type HANDOVER to continue, or N to stop' @('HANDOVER', 'N', 'n')) -ne 'HANDOVER') { return }

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
        icacls "$($old.Path)" /setowner "$sidRef" /T /C /Q 2>&1 | Add-Content $log
        icacls "$($old.Path)" /grant "${sidRef}:(OI)(CI)F" /T /C /Q 2>&1 | Add-Content $log
        # 3. Registry hives inside the profile
        Grant-HiveAccess (Join-Path $old.Path 'NTUSER.DAT') $new.Sid
        $usr = Join-Path $old.Path 'AppData\Local\Microsoft\Windows\UsrClass.dat'
        if (Test-Path $usr) { Grant-HiveAccess $usr $new.Sid }
        # 4. Verify
        $check = Get-CimInstance Win32_UserProfile | Where-Object SID -eq $new.Sid
        if ($check.LocalPath -ne $old.Path) { throw "Verification failed: profile for $($new.Sid) is $($check.LocalPath)" }
        Write-Ok "Done. $($new.Account) now uses $($old.Path). Empty profile parked at $parked (delete it after 30 days)."
        Write-Host 'Next: the person signs in. OneDrive, Outlook and Teams will ask for their new sign-in once; files, desktop and settings are already there.' -ForegroundColor Cyan
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
   0  Exit
"@ -ForegroundColor Cyan
    switch (Read-Choice 'Choose' @('0', '1', '2', '3', '4', '5', '6')) {
        '1' { $Action = 'BuildPackage' }
        '2' { $Action = 'CopyIntune' }
        '3' { $Action = if ((Read-Choice 'Capture this device (C) or move captured hashes between tenants (M)?' @('C', 'M', 'c', 'm')) -in 'C', 'c') { 'AutopilotCapture' } else { 'AutopilotMove' } }
        '4' { $Action = 'MoveDevice' }
        '5' { $Action = 'HandoverProfile' }
        '6' { $Action = 'CopyProfile' }
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
}
