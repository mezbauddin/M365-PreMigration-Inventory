<#
.SYNOPSIS
    Lists every device registered in a tenant's Entra ID, with join type, OS, owner, last sign-in and Intune state,
    so you know which laptops have to be rejoined to the destination and how long it will take.

.DESCRIPTION
    TL;DR: Read-only. Run it against the SOURCE before you plan device rejoins. It reads Entra ID devices through
    Microsoft Graph (trustType AzureAd = Entra joined, ServerAd = hybrid joined, Workplace = registered), their
    registered owners and users, approximate last sign-in, and, unless -SkipIntune, the Intune managed device that
    matches each one (compliance, last sync, serial number, model). Every row gets an Action: Rejoin (Entra joined),
    Re-register (registered phones and personal PCs), On-prem AD: unaffected (hybrid joined), or Stale (not seen in
    -StaleDays). Writes DeviceInventory-<tenant>-<date>.csv.

    Tip: run it against the destination after the rejoins too, and compare. Every source Rejoin row should have a
    matching Entra joined device in the destination.
    Trap: approximateLastSignInDateTime is approximate (updated about every 14 days). Use it to find dead devices,
    not to prove a device is in use today.

.PARAMETER TenantId
    Tenant ID (GUID) to read. The script stops if Graph is signed in anywhere else.

.PARAMETER AdminUpn
    Optional admin sign-in name, e.g. admin@fabrikam.onmicrosoft.com. Any Graph session for another account or tenant
    is disconnected first, and the script prints who and where before it reads anything.

.PARAMETER StaleDays
    Devices with no sign-in for this many days are marked Stale. Default 90.

.PARAMETER SkipIntune
    Don't read Intune managed devices (no DeviceManagementManagedDevices.Read.All consent needed).

.PARAMETER OutputFolder
    Where to write the CSV. Default: current folder.

.EXAMPLE
    .\Get-DeviceInventory.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com

.EXAMPLE
    .\Get-DeviceInventory.ps1 -TenantId <dest-guid> -AdminUpn admin@contoso.onmicrosoft.com -StaleDays 30

.EXAMPLE
    .\Get-DeviceInventory.ps1 -TenantId <source-guid> -SkipIntune

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.0.0
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: Microsoft.Graph.Authentication. Graph scopes Device.Read.All and User.Read.All, plus
              DeviceManagementManagedDevices.Read.All for Intune. Global Reader is enough to read.
    Changes nothing in any tenant. The CSV lists people and their devices: keep it out of source control.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F-]{36}$')][string]$TenantId,
    [string]$AdminUpn,
    [ValidateRange(1, 3650)][int]$StaleDays = 90,
    [switch]$SkipIntune,
    [string]$OutputFolder = '.'
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

function Get-GraphAll([string]$Uri) {
    # Follows @odata.nextLink so tenants with more than one page of devices come back whole.
    $out = [System.Collections.Generic.List[object]]::new()
    while ($Uri) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $Uri
        foreach ($v in @($page.value)) { $out.Add($v) }
        $Uri = $page.'@odata.nextLink'
    }
    $out
}

$scopes = @('Device.Read.All', 'User.Read.All')
if (-not $SkipIntune) { $scopes += 'DeviceManagementManagedDevices.Read.All' }
$ctx = Connect-GraphTenant $TenantId $AdminUpn $scopes
Write-Host "Signed in as $($ctx.Account) to tenant $($ctx.TenantId)" -ForegroundColor Cyan
$tenantName = ((Invoke-MgGraphRequest GET 'v1.0/domains').value | Where-Object { $_.isInitial }).id

Write-Host 'Reading Entra ID devices' -ForegroundColor Cyan
$select = 'id,deviceId,displayName,operatingSystem,operatingSystemVersion,trustType,accountEnabled,isManaged,isCompliant,approximateLastSignInDateTime,registrationDateTime,profileType'
$devices = @(Get-GraphAll "v1.0/devices?`$select=$select&`$expand=registeredOwners")
Write-Host "  $($devices.Count) device(s)" -ForegroundColor Cyan

$intune = @{}
if (-not $SkipIntune) {
    Write-Host 'Reading Intune managed devices' -ForegroundColor Cyan
    try {
        $sel = 'id,deviceName,azureADDeviceId,userPrincipalName,managementAgent,complianceState,lastSyncDateTime,enrolledDateTime,serialNumber,model,manufacturer'
        foreach ($m in @(Get-GraphAll "v1.0/deviceManagement/managedDevices?`$select=$sel")) {
            if ($m.azureADDeviceId -and $m.azureADDeviceId -ne '00000000-0000-0000-0000-000000000000') { $intune["$($m.azureADDeviceId)".ToLower()] = $m }
        }
        Write-Host "  $($intune.Count) Intune device(s) matched to Entra ID" -ForegroundColor Cyan
    }
    catch { Write-Warning "Couldn't read Intune ($($_.Exception.Message)). Carrying on without it. Use -SkipIntune to stop asking." }
}

$joinName = @{ AzureAd = 'Entra joined'; ServerAd = 'Hybrid joined'; Workplace = 'Registered' }
$cutoff = (Get-Date).ToUniversalTime().AddDays(-$StaleDays)
$rows = foreach ($d in $devices) {
    $users = @()
    try { $users = @((Get-GraphAll "v1.0/devices/$($d.id)/registeredUsers?`$select=userPrincipalName") | ForEach-Object { $_.userPrincipalName } | Where-Object { $_ }) }
    catch { $users = @() }
    $owners = @(@($d.registeredOwners) | ForEach-Object { $_.userPrincipalName } | Where-Object { $_ })
    $last = if ($d.approximateLastSignInDateTime) { [datetime]$d.approximateLastSignInDateTime } else { $null }
    $im = $intune["$($d.deviceId)".ToLower()]
    $trust = "$($d.trustType)"
    $os = "$($d.operatingSystem)"
    $action = switch ($trust) {
        'AzureAd'   { 'Rejoin: leave the source, join the destination' }
        'ServerAd'  { 'On-prem AD: unaffected. Hybrid registration follows Entra Connect' }
        'Workplace' { if ($os -match 'iOS|iPhone|iPad|Android') { 'Re-register: remove and re-add the work account' } else { 'Re-register: Access work or school, disconnect, connect' } }
        default     { 'Check: unknown join type' }
    }
    if (-not $last -or $last.ToUniversalTime() -lt $cutoff) { $action = "Stale (no sign-in in $StaleDays days): confirm before you spend time on it. $action" }
    if ($d.accountEnabled -eq $false) { $action = "Disabled in Entra ID. $action" }
    [pscustomobject]@{
        DisplayName     = $d.displayName
        JoinType        = $(if ($joinName[$trust]) { $joinName[$trust] } else { $trust })
        TrustType       = $trust
        OS              = $os
        OSVersion       = $d.operatingSystemVersion
        Owner           = ($owners -join ';')
        RegisteredUsers = ($users -join ';')
        LastSignIn      = $(if ($last) { $last.ToString('yyyy-MM-dd') } else { '' })
        Registered      = $(if ($d.registrationDateTime) { ([datetime]$d.registrationDateTime).ToString('yyyy-MM-dd') } else { '' })
        Enabled         = $d.accountEnabled
        EntraManaged    = $d.isManaged
        EntraCompliant  = $d.isCompliant
        IntuneManaged   = [bool]$im
        IntuneUser      = $(if ($im) { $im.userPrincipalName } else { '' })
        Compliance      = $(if ($im) { $im.complianceState } else { '' })
        IntuneLastSync  = $(if ($im -and $im.lastSyncDateTime) { ([datetime]$im.lastSyncDateTime).ToString('yyyy-MM-dd') } else { '' })
        SerialNumber    = $(if ($im) { $im.serialNumber } else { '' })
        Model           = $(if ($im) { "$($im.manufacturer) $($im.model)".Trim() } else { '' })
        DeviceId        = $d.deviceId
        Action          = $action
    }
}
$rows = @($rows | Sort-Object JoinType, Owner, DisplayName)

$file = Join-Path $OutputFolder "DeviceInventory-$tenantName-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
$rows | Export-Csv $file -NoTypeInformation
Write-Host ''
$rows | Group-Object JoinType | Sort-Object Name | ForEach-Object { Write-Host ("{0,-15} {1}" -f $_.Name, $_.Count) }
$rejoin = @($rows | Where-Object { $_.TrustType -eq 'AzureAd' -and $_.Action -notlike 'Stale*' -and $_.Action -notlike 'Disabled*' })
Write-Host ("{0} Entra joined device(s) in use to rejoin. At 45 to 60 minutes each, allow about {1} to {2} hours of hands-on time." -f
    $rejoin.Count, [math]::Round($rejoin.Count * 0.75, 1), $rejoin.Count) -ForegroundColor Yellow
Write-Host "Report: $file" -ForegroundColor Cyan
Write-Host 'Next: Export-BitLockerKeys.ps1 against the same tenant, before any device leaves it.' -ForegroundColor Cyan
