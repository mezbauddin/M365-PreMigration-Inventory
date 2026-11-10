<#
.SYNOPSIS
    After cutover, blocks sign-in on the source accounts so nobody keeps working in the old tenant, and proves each
    block took by reading it back.

.DESCRIPTION
    TL;DR: Once people are working in the destination, their source accounts are a trap: Outlook on a phone, a saved
    browser session or a shared PC keeps them sending and receiving in the old tenant, and that mail never arrives
    in the new one. This script sets accountEnabled to false for each source account in the mapping file (or a list
    you give it), reads the account back to check it really is blocked, and writes a CSV. Your own admin account is
    always skipped. Supports -WhatIf.

    Accounts are found by UPN, then by mail or proxy address, then (because Remove-SourceDomain moved everyone to
    onmicrosoft.com) by <local part>@<tenant>.onmicrosoft.com.

    Trap: accounts that hold an admin role can't be blocked by a User Administrator, and often not through Graph at
    all: Graph returns 403 Authorization_RequestDenied. The script says which role, and you block those in the
    Microsoft 365 admin centre (https://admin.microsoft.com/#/users > the user > Block sign-in).
    Trap: an earlier version of a loop like this printed "Blocked" even when the update failed. This one only says
    Blocked after reading accountEnabled back as false.
    Tip: blocking doesn't end sessions that are already signed in. Revoke them too (Revoke-MgUserSignInSession, or
    "Sign out of all sessions" in the admin centre) if you need it immediate.
    Tip: keep one break-glass Global Administrator in the source unblocked until it is decommissioned.
    Paid tool: not needed. This is a one-property change; the value is in checking it took.
    No script: Microsoft 365 admin centre > Active users > select them > Block sign-in. Fine for a handful.

.PARAMETER TenantId
    Source tenant ID (GUID). The script stops if Graph is signed in anywhere else.

.PARAMETER AdminUpn
    Source admin account. Any Graph session for another account or tenant is disconnected first. Never blocked.

.PARAMETER MappingFile
    The mapping CSV (Source, Destination). The Source column is used.

.PARAMETER User
    Instead of a mapping file, the source UPNs or addresses to block.

.PARAMETER Exclude
    UPNs or addresses never to block, e.g. a break-glass admin or a service account still in use.

.EXAMPLE
    .\Block-SourceSignIn.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -MappingFile .\UserMapping.csv -WhatIf

.EXAMPLE
    .\Block-SourceSignIn.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -MappingFile .\UserMapping.csv -Exclude breakglass@fabrikam.onmicrosoft.com

.EXAMPLE
    .\Block-SourceSignIn.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -User alex@fabrikam.com, sam@fabrikam.com

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.0.0
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: Microsoft.Graph.Authentication, Microsoft.Graph.Users. User Administrator (Privileged Authentication
              Administrator or Global Administrator for admin accounts) in the source.
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Mapping')]
param(
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F-]{36}$')][string]$TenantId,
    [Parameter(Mandatory)][string]$AdminUpn,
    [Parameter(Mandatory, ParameterSetName = 'Mapping')][string]$MappingFile,
    [Parameter(Mandatory, ParameterSetName = 'List')][string[]]$User,
    [string[]]$Exclude = @()
)

$ErrorActionPreference = 'Stop'

function Connect-GraphTenant([string]$TenantId, [string]$AdminUpn, [string[]]$Scopes) {
    $WhatIfPreference = $false
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

$targets = if ($PSCmdlet.ParameterSetName -eq 'Mapping') {
    if (-not (Test-Path $MappingFile)) { throw "Mapping file $MappingFile not found." }
    @(Import-Csv $MappingFile | Where-Object Source | ForEach-Object { $_.Source.Trim().ToLower() })
}
else { @($User | ForEach-Object { $_.Trim().ToLower() }) }
$targets = @($targets | Where-Object { $_ } | Select-Object -Unique)
$skip = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
@($Exclude) + $AdminUpn | Where-Object { $_ } | ForEach-Object { [void]$skip.Add($_.Trim()) }

$ctx = Connect-GraphTenant $TenantId $AdminUpn @('User.ReadWrite.All', 'Directory.Read.All')
Write-Host "Signed in as $($ctx.Account) to tenant $($ctx.TenantId)" -ForegroundColor Cyan
$initial = ((Invoke-MgGraphRequest GET 'v1.0/domains').value | Where-Object { $_.isInitial }).id
Write-Host "$($targets.Count) account(s) to block. Fallback domain: $initial" -ForegroundColor Cyan

function Find-SourceUser([string]$Address) {
    $props = 'id,userPrincipalName,displayName,accountEnabled'
    $q = $Address.Replace("'", "''")
    $local = $Address.Split('@')[0].Replace("'", "''")
    foreach ($f in "userPrincipalName eq '$q'", "mail eq '$q'", "proxyAddresses/any(p:p eq 'smtp:$q')", "userPrincipalName eq '$local@$initial'") {
        $r = @((Invoke-MgGraphRequest GET "v1.0/users?`$filter=$([uri]::EscapeDataString($f))&`$select=$props").value)
        if ($r.Count -eq 1) { return $r[0] }
    }
    $null
}
function Get-AdminRoles([string]$Id) {
    try {
        @((Invoke-MgGraphRequest GET "v1.0/users/$Id/memberOf/microsoft.graph.directoryRole?`$select=displayName").value | ForEach-Object { $_.displayName })
    }
    catch { @() }
}

$log = [System.Collections.Generic.List[object]]::new()
function Write-Log($Source, $Upn, $Status, $Detail) {
    $log.Add([pscustomobject]@{ Source = $Source; UserPrincipalName = $Upn; Status = $Status; Detail = $Detail })
    $c = @{ Blocked = 'Green'; AlreadyBlocked = 'Gray'; WhatIf = 'Cyan'; Skipped = 'Gray'; NotFound = 'Yellow'; AdminRole = 'Magenta'; Problem = 'Red' }[$Status]
    Write-Host ("  [{0}] {1} {2} {3}" -f $Status, $Source, $(if ($Upn -and $Upn -ne $Source) { "($Upn)" }), $Detail) -ForegroundColor $c
}

foreach ($t in $targets) {
    if ($skip.Contains($t)) { Write-Log $t '' 'Skipped' 'Excluded'; continue }
    $u = Find-SourceUser $t
    if (-not $u) { Write-Log $t '' 'NotFound' 'No single source account matches'; continue }
    $upn = $u.userPrincipalName
    if ($skip.Contains($upn)) { Write-Log $t $upn 'Skipped' 'Excluded (or your own admin account)'; continue }
    if ($u.accountEnabled -eq $false) { Write-Log $t $upn 'AlreadyBlocked' ''; continue }
    if (-not $PSCmdlet.ShouldProcess($upn, 'Block sign-in')) { Write-Log $t $upn 'WhatIf' 'Would block sign-in'; continue }
    try {
        Invoke-MgGraphRequest PATCH "v1.0/users/$($u.id)" -ContentType 'application/json' -Body (@{ accountEnabled = $false } | ConvertTo-Json) | Out-Null
        $check = Invoke-MgGraphRequest GET "v1.0/users/$($u.id)?`$select=accountEnabled"
        if ($check.accountEnabled -eq $false) { Write-Log $t $upn 'Blocked' 'Read back: accountEnabled false' }
        else { Write-Log $t $upn 'Problem' 'Update returned, but the account still reads as enabled. Check in the admin centre' }
    }
    catch {
        $msg = "$($_.Exception.Message) $($_.ErrorDetails.Message)"
        if ($msg -match 'Authorization_RequestDenied|Forbidden|403|Insufficient privileges') {
            $roles = @(Get-AdminRoles $u.id)
            if ($roles) { Write-Log $t $upn 'AdminRole' "Holds an admin role ($($roles -join ', ')): block it in the admin centre: https://admin.microsoft.com/#/users" }
            else { Write-Log $t $upn 'Problem' "Access denied. Your account may need User Administrator or higher: $($_.Exception.Message)" }
        }
        else { Write-Log $t $upn 'Problem' $_.Exception.Message }
    }
}

$file = ".\BlockSourceSignIn-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
$log | Export-Csv $file -NoTypeInformation -WhatIf:$false
Write-Host ''
$log | Group-Object Status | Sort-Object Name | ForEach-Object { Write-Host ("{0,-15} {1}" -f $_.Name, $_.Count) }
Write-Host "Report: $file" -ForegroundColor Cyan
if ($log | Where-Object Status -eq 'AdminRole') { Write-Host 'Admin accounts above: block them by hand in the Microsoft 365 admin centre.' -ForegroundColor Magenta }
