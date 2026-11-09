<#
.SYNOPSIS
    Frees a domain held by deleted users in the recycle bin: restores each one under an onmicrosoft.com sign-in,
    blocks it, and deletes it again. Nothing is purged.

.DESCRIPTION
    TL;DR: Microsoft won't remove a domain from a tenant while any deleted user in the recycle bin still has a sign-in
    on it, and it won't rename a user while it's deleted. This script finds those users, restores each one as
    <name>-old@<tenant>.onmicrosoft.com, blocks sign-in, and puts it back in the recycle bin. The domain is free and
    every account can still be recovered for another 30 days. Use -WhatIf to see the list first.

.PARAMETER TenantId
    Source tenant ID. The script stops if you're signed in somewhere else.

.PARAMETER Domain
    The domains you're about to remove.

.EXAMPLE
    .\Clear-DeletedUserDomain.ps1 -TenantId <source-guid> -Domain fabrikam.com, fabrikam.co.uk -WhatIf

.EXAMPLE
    .\Clear-DeletedUserDomain.ps1 -TenantId <source-guid> -Domain fabrikam.com, fabrikam.co.uk

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.0.0
    Date    : 2026-11-09
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: Microsoft.Graph.Authentication, User Administrator (or higher) in the source tenant.
    Run it in a fresh PowerShell window if other Graph or Exchange scripts ran there first.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F-]{36}$')][string]$TenantId,
    [Parameter(Mandatory)][string[]]$Domain
)

$ErrorActionPreference = 'Stop'
Connect-MgGraph -TenantId $TenantId -Scopes User.ReadWrite.All, Domain.Read.All -NoWelcome
if ((Get-MgContext).TenantId -ne $TenantId) { throw "Signed in to $((Get-MgContext).TenantId), not $TenantId. Run Disconnect-MgGraph and try again." }

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
    try {
        Invoke-MgGraphRequest POST "v1.0/directory/deletedItems/$($u.id)/restore" -ContentType 'application/json' `
            -Body (@{ newUserPrincipalName = $new; autoReconcileProxyConflict = $true } | ConvertTo-Json) | Out-Null
        Invoke-MgGraphRequest PATCH "v1.0/users/$($u.id)" -ContentType 'application/json' -Body (@{ accountEnabled = $false } | ConvertTo-Json)
        Invoke-MgGraphRequest DELETE "v1.0/users/$($u.id)"
        Write-Host "  OK   $name -> $new (back in the recycle bin)" -ForegroundColor Green
        [pscustomobject]@{ User = $u.userPrincipalName; NewSignIn = $new; Result = 'Done'; Detail = '' }
    }
    catch {
        Write-Host "  FAIL $name : $($_.Exception.Message)" -ForegroundColor Red
        [pscustomobject]@{ User = $u.userPrincipalName; NewSignIn = $new; Result = 'Problem'; Detail = $_.Exception.Message }
    }
}
$file = ".\ClearDeletedUserDomain-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
$log | Export-Csv $file -NoTypeInformation -WhatIf:$false
Write-Host "Report: $file" -ForegroundColor Cyan
