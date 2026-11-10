<#
.SYNOPSIS
    Frees a domain held by deleted users in the recycle bin: restores each one under an onmicrosoft.com sign-in,
    blocks it, and deletes it again. Nothing is purged.

.DESCRIPTION
    TL;DR: Microsoft won't remove a domain from a tenant while any deleted user in the recycle bin still has a sign-in
    on it, and it won't rename a user while it's deleted. This script finds those users, restores each one as
    <name>-old@<tenant>.onmicrosoft.com, blocks sign-in, and puts it back in the recycle bin. The domain is free and
    every account can still be recovered for another 30 days. Use -WhatIf to see the list first.

    Tip: run it the day before cutover, not on the night.
    Trap: deleted accounts that held an admin role can't be restored by a User Administrator. Use a Global
    Administrator or Privileged Role Administrator, or purge them if nobody needs them back.
    Paid tool: not needed.
    No script: Entra admin centre > Users > Deleted users > restore one, rename it in Properties, block sign-in, then
    delete it again. Fine for two or three accounts, slow for twenty.

.PARAMETER TenantId
    Source tenant ID. The script stops if you're signed in somewhere else.

.PARAMETER Domain
    The domains you're about to remove.

.PARAMETER AdminUpn
    Optional source admin account. If Graph is signed in as someone else, or to another tenant, it is disconnected
    and the script signs in again; it stops if the account you pick doesn't match.

.EXAMPLE
    .\Clear-DeletedUserDomain.ps1 -TenantId <source-guid> -Domain fabrikam.com, fabrikam.co.uk -WhatIf

.EXAMPLE
    .\Clear-DeletedUserDomain.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -Domain fabrikam.com, fabrikam.co.uk

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.1.0
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: Microsoft.Graph.Authentication, User Administrator (or higher) in the source tenant.
    It disconnects any Graph session that belongs to another tenant or account before it starts.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F-]{36}$')][string]$TenantId,
    [Parameter(Mandatory)][string[]]$Domain,
    [string]$AdminUpn
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
$ctx = Connect-GraphTenant $TenantId $AdminUpn @('User.ReadWrite.All', 'Domain.Read.All')
Write-Host "Signed in as $($ctx.Account) to tenant $($ctx.TenantId)" -ForegroundColor Cyan

$fallback = ((Invoke-MgGraphRequest GET 'v1.0/domains').value | Where-Object { $_.isInitial }).id
$domains = @($Domain | ForEach-Object { $_.ToLower().TrimStart('@') })

$deleted = [System.Collections.Generic.List[object]]::new()
$uri = 'v1.0/directory/deletedItems/microsoft.graph.user?$select=id,userPrincipalName,displayName&$top=999'
while ($uri) { $page = Invoke-MgGraphRequest GET $uri; $page.value | ForEach-Object { $deleted.Add($_) }; $uri = $page.'@odata.nextLink' }

$hits = @($deleted | Where-Object { $domains -contains "$($_.userPrincipalName)".Split('@')[-1].ToLower() })
if (-not $hits) { Write-Host "No deleted users hold $($domains -join ', '). Nothing to do." -ForegroundColor Green; return }
Write-Host "$($hits.Count) deleted user(s) hold $($domains -join ', '). New sign-ins go on $fallback." -ForegroundColor Cyan

$log = foreach ($u in $hits) {
    # Deleted users' UPNs are prefixed with their object ID without dashes
    $name = ($u.userPrincipalName -replace '^[0-9a-f]{32}', '').Split('@')[0]
    $new = "$name-old@$fallback"
    if (-not $PSCmdlet.ShouldProcess("$($u.displayName) ($name)", "Restore as $new, block, delete again")) {
        [pscustomobject]@{ User = $u.userPrincipalName; NewSignIn = $new; Result = 'WhatIf'; Detail = '' }; continue
    }
    $step = 'restore'
    try {
        Invoke-MgGraphRequest POST "v1.0/directory/deletedItems/$($u.id)/restore" -ContentType 'application/json' `
            -Body (@{ newUserPrincipalName = $new; autoReconcileProxyConflict = $true } | ConvertTo-Json) | Out-Null
        $step = 'block sign-in'
        Invoke-MgGraphRequest PATCH "v1.0/users/$($u.id)" -ContentType 'application/json' -Body (@{ accountEnabled = $false } | ConvertTo-Json) | Out-Null
        $step = 'delete again'
        Invoke-MgGraphRequest DELETE "v1.0/users/$($u.id)" | Out-Null
        Write-Host "  OK   $name -> $new (back in the recycle bin)" -ForegroundColor Green
        [pscustomobject]@{ User = $u.userPrincipalName; NewSignIn = $new; Result = 'Done'; Detail = '' }
    }
    catch {
        $why = if ("$($_.Exception.Message)" -match 'Forbidden' -and $step -eq 'restore') { ' Usually the account held an admin role: restoring it needs Privileged Role Administrator or Global Administrator.' } else { '' }
        Write-Host "  FAIL $name at '$step': $($_.Exception.Message)$why" -ForegroundColor Red
        [pscustomobject]@{ User = $u.userPrincipalName; NewSignIn = $new; Result = 'Problem'; Detail = "$step : $($_.Exception.Message)$why" }
    }
}
$file = ".\ClearDeletedUserDomain-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
@($log | Where-Object { $_ }) | Export-Csv $file -NoTypeInformation -WhatIf:$false
Write-Host "Report: $file" -ForegroundColor Cyan
