<#
.SYNOPSIS
    Exports every BitLocker recovery key stored in a tenant's Entra ID, with the device name each one belongs to, so
    the keys survive the source tenant being closed.

.DESCRIPTION
    TL;DR: Run it against the SOURCE before any device leaves it. Recovery keys for Entra joined devices are kept in
    the tenant the device was joined to. Once the device rejoins the destination, or the source is closed, those keys
    are gone, and a laptop that asks for its recovery key after a BIOS update or a new motherboard is a laptop you
    can't open. This script lists every key through Microsoft Graph (informationProtection/bitlocker/recoveryKeys),
    reads each key value, maps its deviceId to the device name, owner and OS, and writes one CSV into a folder that
    only you can read. Read-only in the tenant.

    Trap: the CSV holds the keys to every disk in the company. Treat it like a password list. Don't put it in
    OneDrive, SharePoint, email or a Git repo. Put it in your password manager or an encrypted store the business
    controls, then delete the local copy.
    Trap: every key read is audited in Entra ID (KeyManagement). Tell the security team before you run it, so a
    burst of "Read BitLocker key" events isn't a surprise.
    Tip: after each device rejoins the destination, it backs up a new key there (if your Intune or BitLocker policy
    asks it to). Check it has, then the old key is only needed for old backups of that disk.

.PARAMETER TenantId
    Tenant ID (GUID) to read. The script stops if Graph is signed in anywhere else.

.PARAMETER AdminUpn
    Optional admin sign-in name, e.g. admin@fabrikam.onmicrosoft.com. Any Graph session for another account or tenant
    is disconnected first, and the script prints who and where before it reads anything.

.PARAMETER OutputFolder
    Folder for the CSV. Default: .\BitLockerKeys-<tenant>-<date>. Created if missing. The script removes inherited
    permissions and leaves only the current user with access (chmod 700 on macOS and Linux).

.PARAMETER Force
    Write even if the folder looks like it's synced to OneDrive or inside a Git repository.

.EXAMPLE
    .\Export-BitLockerKeys.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com

.EXAMPLE
    .\Export-BitLockerKeys.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -OutputFolder D:\Secure\BitLocker

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.0.0
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: Microsoft.Graph.Authentication. Graph scopes BitLockerKey.Read.All and Device.Read.All. The signed-in
              account needs a role that can read keys: Cloud Device Administrator, Helpdesk Administrator, Intune
              Administrator, Security Administrator, Security Reader or Global Reader.
    Changes nothing in the tenant. The output is the most sensitive file in this toolkit.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F-]{36}$')][string]$TenantId,
    [string]$AdminUpn,
    [string]$OutputFolder,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) { throw 'Install-Module Microsoft.Graph.Authentication -Scope CurrentUser first.' }

function Connect-GraphTenant([string]$TenantId, [string]$AdminUpn, [string[]]$Scopes) {
    $ctx = Get-MgContext
    if ($ctx -and ($ctx.TenantId -ne $TenantId -or ($AdminUpn -and $ctx.Account -ne $AdminUpn))) {
        Write-Host "Graph is signed in as $($ctx.Account) to $($ctx.TenantId). Disconnecting." -ForegroundColor Yellow
        Disconnect-MgGraph | Out-Null
    }
    Connect-MgGraph -TenantId $TenantId -Scopes $Scopes -NoWelcome
    $ctx = Get-MgContext
    if ($ctx.TenantId -ne $TenantId) { throw "Graph signed in to $($ctx.TenantId), not $TenantId. Stopping." }
    if ($AdminUpn -and $ctx.Account -ne $AdminUpn) { throw "Graph signed in as $($ctx.Account), not $AdminUpn. Run Disconnect-MgGraph and pick $AdminUpn." }
    $ctx
}

function Get-GraphAll([string]$Uri, [hashtable]$Headers = @{}) {
    $out = [System.Collections.Generic.List[object]]::new()
    while ($Uri) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $Uri -Headers $Headers
        foreach ($v in @($page.value)) { $out.Add($v) }
        $Uri = $page.'@odata.nextLink'
    }
    $out
}

function Test-UnsafeFolder([string]$Path) {
    # Synced or versioned folders copy the keys somewhere you didn't mean them to go.
    $full = [IO.Path]::GetFullPath($Path)
    if ($full -match '(?i)onedrive|sharepoint|dropbox|google drive|icloud') { return "it looks like a synced folder ($full)" }
    $dir = [IO.DirectoryInfo]::new($full)
    while ($dir) {
        if (Test-Path (Join-Path $dir.FullName '.git')) { return "it's inside a Git repository ($($dir.FullName))" }
        $dir = $dir.Parent
    }
    $null
}

$ctx = Connect-GraphTenant $TenantId $AdminUpn @('BitLockerKey.Read.All', 'Device.Read.All')
Write-Host "Signed in as $($ctx.Account) to tenant $($ctx.TenantId)" -ForegroundColor Cyan
$tenantName = ((Invoke-MgGraphRequest GET 'v1.0/domains').value | Where-Object { $_.isInitial }).id

if (-not $OutputFolder) { $OutputFolder = Join-Path (Get-Location) "BitLockerKeys-$tenantName-$(Get-Date -Format 'yyyyMMdd-HHmm')" }
$unsafe = Test-UnsafeFolder $OutputFolder
if ($unsafe -and -not $Force) { throw "Not writing BitLocker keys to $OutputFolder because $unsafe. Pick a local, unsynced folder with -OutputFolder, or use -Force if you're sure." }
if ($unsafe) { Write-Warning "Writing anyway (-Force), although $unsafe." }
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null

# Lock the folder down to the current user on Windows. Elsewhere, set permissions yourself.
if ($IsWindows -or $PSVersionTable.PSEdition -eq 'Desktop') {
    try {
        $me = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $acl = Get-Acl $OutputFolder
        $acl.SetAccessRuleProtection($true, $false)
        foreach ($rule in @($acl.Access)) { [void]$acl.RemoveAccessRule($rule) }
        $acl.AddAccessRule([System.Security.AccessControl.FileSystemAccessRule]::new($me, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow'))
        Set-Acl -Path $OutputFolder -AclObject $acl
        Write-Host "Folder permissions: $me only." -ForegroundColor Cyan
    }
    catch { Write-Warning "Couldn't restrict folder permissions: $($_.Exception.Message). Restrict them yourself before you go on." }
}
else {
    try { & chmod 700 $OutputFolder; Write-Host 'Folder permissions: owner only (chmod 700).' -ForegroundColor Cyan }
    catch { Write-Warning "Couldn't restrict folder permissions. Run chmod 700 on $OutputFolder yourself." }
}

# User-Agent is required by the recovery key API; ocp-client-* are optional but help Microsoft support.
$headers = @{ 'User-Agent' = 'M365-T2T-Toolkit/1.0'; 'ocp-client-name' = 'M365 T2T Toolkit Export-BitLockerKeys'; 'ocp-client-version' = '1.0.0' }

Write-Host 'Reading devices' -ForegroundColor Cyan
$devices = @{}
foreach ($d in @(Get-GraphAll "v1.0/devices?`$select=id,deviceId,displayName,operatingSystem,trustType&`$expand=registeredOwners")) {
    if ($d.deviceId) { $devices["$($d.deviceId)".ToLower()] = $d }
}
Write-Host "  $($devices.Count) device(s)" -ForegroundColor Cyan

Write-Host 'Listing recovery keys' -ForegroundColor Cyan
$keys = @(Get-GraphAll 'v1.0/informationProtection/bitlocker/recoveryKeys' $headers)
Write-Host "  $($keys.Count) key(s). Reading each one (each read is audited)." -ForegroundColor Cyan

$rows = [System.Collections.Generic.List[object]]::new()
$failed = 0
$i = 0
foreach ($k in $keys) {
    $i++; Write-Progress -Activity 'BitLocker keys' -Status "$i of $($keys.Count)" -PercentComplete (100 * $i / [math]::Max(1, $keys.Count))
    $value = $null; $note = ''
    try { $value = (Invoke-MgGraphRequest -Method GET -Uri "v1.0/informationProtection/bitlocker/recoveryKeys/$($k.id)?`$select=key" -Headers $headers).key }
    catch { $failed++; $note = "Couldn't read the key: $($_.Exception.Message)" }
    $d = $devices["$($k.deviceId)".ToLower()]
    $rows.Add([pscustomobject]@{
            DeviceName  = $(if ($d) { $d.displayName } else { '' })
            Owner       = $(if ($d) { (@($d.registeredOwners) | ForEach-Object { $_.userPrincipalName } | Where-Object { $_ }) -join ';' } else { '' })
            OS          = $(if ($d) { $d.operatingSystem } else { '' })
            JoinType    = $(if ($d) { $d.trustType } else { '' })
            VolumeType  = $k.volumeType
            KeyId       = $k.id
            RecoveryKey = $value
            Created     = $(if ($k.createdDateTime) { ([datetime]$k.createdDateTime).ToString('yyyy-MM-dd HH:mm') } else { '' })
            DeviceId    = $k.deviceId
            Note        = $(if (-not $d -and -not $note) { 'Device no longer in Entra ID. Keep the key anyway: the disk may still exist' } else { $note })
        })
}
Write-Progress -Activity 'BitLocker keys' -Completed

$file = Join-Path $OutputFolder "BitLockerKeys-$tenantName-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
$rows | Sort-Object DeviceName, Created | Export-Csv $file -NoTypeInformation
$withKey = @($rows | Where-Object RecoveryKey).Count
$noKeyDevices = @($devices.Values | Where-Object { $_.trustType -eq 'AzureAd' -and "$($_.operatingSystem)" -match 'Windows' -and
        ($keys.deviceId | ForEach-Object { "$_".ToLower() }) -notcontains "$($_.deviceId)".ToLower() })
Write-Host ''
Write-Host "$withKey key(s) exported for $(@($rows | Where-Object RecoveryKey | Select-Object -ExpandProperty DeviceId -Unique).Count) device(s)." -ForegroundColor Green
if ($failed) { Write-Host "$failed key(s) couldn't be read. Check your role (see NOTES) and the Note column." -ForegroundColor Red }
if ($noKeyDevices) { Write-Host "$($noKeyDevices.Count) Entra joined Windows device(s) have no key in Entra ID: $(($noKeyDevices.displayName | Select-Object -First 10) -join ', '). Check BitLocker on those before they leave (manage-bde -protectors -get C:)." -ForegroundColor Yellow }
Write-Warning "$file holds every recovery key. Move it into your password manager or an encrypted store, then delete this copy."
