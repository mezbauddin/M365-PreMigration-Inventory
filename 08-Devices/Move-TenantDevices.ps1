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
    Version : 1.0.1
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: Windows PowerShell 5.1 or 7, elevated. Admin-PC options need Microsoft.Graph.Authentication.
              Laptop options need nothing extra (Provisioning module is built in).
    Status  : Option 1's scripted token request uses the same sign-in flow Windows Configuration Designer uses, which
              Microsoft doesn't document. If it stops working, the script opens WCD for you instead. Options 2, 3 and 5
              are tested against stubbed tenants and devices; run them with -WhatIf first on anything that matters.
    Changelog:
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
function Invoke-BuildPackage {
    Write-Step 'Build the enrolment package (run on your admin PC)'
    if (-not $DestinationTenantId) { $DestinationTenantId = Read-Host 'Destination tenant ID (GUID)' }
    if (-not $DestinationAdmin) { $DestinationAdmin = Read-Host 'Destination admin sign-in (e.g. admin@contoso.com)' }
    $name = Read-Host 'Package name (letters and digits, e.g. contoso-join)'
    if (-not $name) { $name = 'tenant-join' }
    $expiry = Read-Host 'Token validity in days (max 180, default 180)'
    if (-not $expiry) { $expiry = 180 }
    $ppkg = Join-Path $WorkFolder "$name.ppkg"

    # The bulk token is a refresh token issued to the Windows Configuration Designer app for the device registration
    # service. WCD requests it with the same OAuth device-code flow; the account "package_<guid>" that appears in the
    # tenant afterwards is created by the service, not by us.
    $wcdApp = 'b90d5b8f-5503-4153-b545-b31cecfaece2'
    $drs = 'urn:ms-drs:enterpriseregistration.windows.net'
    $bprt = $null
    Write-Host 'Trying the scripted token request (same flow as Windows Configuration Designer)...' -ForegroundColor Cyan
    try {
        $dc = Invoke-RestMethod -Method Post "https://login.microsoftonline.com/$DestinationTenantId/oauth2/devicecode" -Body @{ client_id = $wcdApp; resource = $drs }
        Write-Host $dc.message -ForegroundColor Yellow
        Write-Host "Sign in as $DestinationAdmin. Waiting..." -ForegroundColor Yellow
        $deadline = (Get-Date).AddSeconds([int]$dc.expires_in)
        while ((Get-Date) -lt $deadline) {
            Start-Sleep -Seconds ([int]$dc.interval)
            try {
                $tok = Invoke-RestMethod -Method Post "https://login.microsoftonline.com/$DestinationTenantId/oauth2/token" -Body @{ grant_type = 'device_code'; client_id = $wcdApp; code = $dc.device_code; resource = $drs }
                if ($tok.refresh_token) { $bprt = $tok.refresh_token; break }
            }
            catch {
                $body = ''; try { $body = $_.ErrorDetails.Message } catch {}
                if ($body -match 'authorization_pending') { continue }
                throw
            }
        }
    }
    catch { Write-Warn2 "Scripted token request failed: $($_.Exception.Message)" }

    $icd = Get-ChildItem "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Imaging and Configuration Designer" -Recurse -Filter ICD.exe -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($bprt -and $icd) {
        Write-Ok 'Token received. Building the package with ICD.exe.'
        $id = [guid]::NewGuid().ToString()
        $xml = @"
<?xml version="1.0" encoding="utf-8"?>
<WindowsCustomizations>
  <PackageConfig xmlns="urn:schemas-Microsoft-com:Windows-ICD-Package-Config.v1.0">
    <ID>{$id}</ID><Name>$name</Name><Version>1.0</Version><OwnerType>OEM</OwnerType><Rank>0</Rank>
  </PackageConfig>
  <Settings xmlns="urn:schemas-microsoft-com:windows-provisioning">
    <Customizations><Common><Accounts><Azure>
      <Authority>https://login.microsoftonline.com/common</Authority>
      <BPRT>$bprt</BPRT>
    </Azure></Accounts></Common></Customizations>
  </Settings>
</WindowsCustomizations>
"@
        $xmlPath = Join-Path $WorkFolder "$name.customizations.xml"
        $xml | Set-Content $xmlPath -Encoding UTF8 -WhatIf:$false
        $store = Get-ChildItem (Split-Path $icd.FullName) -Recurse -Filter 'Microsoft-Desktop-Provisioning.dat' -ErrorAction SilentlyContinue | Select-Object -First 1
        $args = @("/Build-ProvisioningPackage", "/CustomizationXML:`"$xmlPath`"", "/PackagePath:`"$ppkg`"", '+Overwrite')
        if ($store) { $args += "/StoreFile:`"$($store.FullName)`"" }
        & $icd.FullName @args | Out-Null
        Remove-Item $xmlPath -Force -ErrorAction SilentlyContinue
        if (Test-Path $ppkg) { Write-Ok "Package built: $ppkg (keep it private: it joins any device to the tenant for $expiry days)"; Add-State 'Package' $ppkg; return }
        Write-Warn2 'ICD.exe did not produce a package.'
    }
    if ($bprt -and -not $icd) { Write-Warn2 'Token received but Windows Configuration Designer (ICD.exe) is not installed to build the package.' }

    # Fallback: Windows Configuration Designer, guided
    Write-Step 'Fallback: Windows Configuration Designer'
    if (-not $icd) {
        Write-Host 'Install it with one of these, then run this option again:' -ForegroundColor Yellow
        Write-Host '  winget install --id 9NBLGGH4TX22 --source msstore        (Microsoft Store version, quickest)'
        Write-Host '  or the Windows ADK, feature "Configuration Designer": adksetup.exe /features OptionId.ICDConfigurationDesigner /quiet'
        return
    }
    Write-Host @"
1. Open Windows Configuration Designer: Provision desktop devices.
2. Name: $name. Set up device: any name pattern, keep the rest default. Set up network: skip (off).
3. Account management: "Enroll in Azure AD", Bulk token expiry: $expiry days, Get Bulk Token, sign in as $DestinationAdmin.
4. Add applications and certificates: skip. Finish: Create. Note the path it shows.
5. Copy the .ppkg to: $ppkg
"@ -ForegroundColor Yellow
    Start-Process $icd.FullName -ErrorAction SilentlyContinue
    $p = Read-Host 'When the package exists, paste its full path (or Enter to stop here)'
    if ($p -and (Test-Path $p)) { Copy-Item $p $ppkg -Force; Write-Ok "Package ready: $ppkg"; Add-State 'Package' $ppkg }
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
