<#
.SYNOPSIS
    Cutover-eve go/no-go. Read-only checks across both tenants and DNS, with a PASS, WARN or FAIL for each.

.DESCRIPTION
    TL;DR: Run it the day before and again an hour before cutover. It signs in to the source and then the
    destination with the accounts you name, checks everything that stopped or slowed a real cutover, and prints a
    table. GO means nothing FAILed. Changes nothing.

    Source:       deleted users holding the domain, synced users on it, Entra users with proxy addresses on it but
                  no mailbox (ProxyOnly), app registrations with identifier URIs on it.
    Snapshot:     it is the -pre one: it has addresses on the domain and LegacyExchangeDN values.
    Mapping:      the file exists and every Destination resolves in the destination.
    Destination:  conflicts (old addresses already on guests, mail users, contacts or other mailboxes), domain state
                  (not added, unverified with the TXT to publish, verified), Conditional Access policies with
                  location conditions (your test sign-in from home may be blocked).
    DNS:          the authoritative name servers, the NS TTL, and whether the destination's MS= TXT is already on
                  every one of them.
    Modules:      runs ..\tools\Test-ModuleHealth.ps1 if it is there.

    Tip: a WARN is something to read, not to ignore. Most of the surprises on a live cutover would have been WARNs.
    Paid tool: migration tools have pre-flight reports for the data copy. None I've used checks domain blockers.
    No script: Entra admin centre (Deleted users, App registrations), Exchange admin centre (recipients), the
    Microsoft 365 admin centre domain page, and nslookup against each name server. Allow an hour.

.PARAMETER SourceTenantId
    Source tenant ID (GUID).

.PARAMETER DestinationTenantId
    Destination tenant ID (GUID).

.PARAMETER SourceAdmin
    Source admin account, e.g. admin@fabrikam.onmicrosoft.com.

.PARAMETER DestinationAdmin
    Destination admin account, e.g. admin@contoso.onmicrosoft.com.

.PARAMETER Domain
    The domains moving, e.g. fabrikam.com.

.PARAMETER MappingFile
    The mapping CSV (Source, Destination).

.PARAMETER SourceSnapshot
    The SourceRecipients-pre-<stamp>.csv from Remove-SourceDomain -SnapshotOnly. Take the snapshot first: at T-1 day
    for the go/no-go, and again on the night for the final run just before the release.

.EXAMPLE
    .\Test-CutoverReadiness.ps1 -SourceTenantId <source-guid> -DestinationTenantId <destination-guid> `
        -SourceAdmin admin@fabrikam.onmicrosoft.com -DestinationAdmin admin@contoso.onmicrosoft.com `
        -Domain fabrikam.com -MappingFile .\UserMapping.csv -SourceSnapshot .\SourceRecipients-pre-<stamp>.csv

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.0.1
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: ExchangeOnlineManagement 3.x, Microsoft.Graph.Authentication. Read roles in both tenants (Global Reader
              plus Exchange recipient read is enough). Windows (Resolve-DnsName) for the DNS checks.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F-]{36}$')][string]$SourceTenantId,
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F-]{36}$')][string]$DestinationTenantId,
    [Parameter(Mandatory)][string]$SourceAdmin,
    [Parameter(Mandatory)][string]$DestinationAdmin,
    [Parameter(Mandatory)][string[]]$Domain,
    [Parameter(Mandatory)][string]$MappingFile,
    [Parameter(Mandatory)][string]$SourceSnapshot
)

$ErrorActionPreference = 'Stop'
$Domain = @($Domain | ForEach-Object { $_.Trim().ToLower() })
function Get-DomainOf($a) { "$a".Split('@')[-1].ToLower() }
function Test-OnDomain($a) { $Domain -contains (Get-DomainOf $a) }
function Test-HostOnDomain($uri) {
    $h = try { ([uri]"$uri").Host.ToLower() } catch { '' }
    if (-not $h) { return $false }
    foreach ($d in $Domain) { if ($h -eq $d -or $h.EndsWith(".$d")) { return $true } }
    $false
}

$results = [System.Collections.Generic.List[object]]::new()
function Add-Result($Area, $Check, $Result, $Detail, $Fix = '') {
    $results.Add([pscustomobject]@{ Area = $Area; Check = $Check; Result = $Result; Detail = $Detail; Fix = $Fix })
    $c = @{ PASS = 'Green'; WARN = 'Yellow'; FAIL = 'Red'; INFO = 'Gray'; SKIP = 'Gray' }[$Result]
    Write-Host ("  {0,-4} {1}: {2}" -f $Result, $Check, $Detail) -ForegroundColor $c
}

# ------------------------------------------------------------ connection helpers (named account, right tenant)
function Connect-ExoTenant([string]$TenantId, [string]$AdminUpn) {
    $c = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object State -eq 'Connected' | Select-Object -First 1
    if ($c -and ($c.TenantID -ne $TenantId -or ($AdminUpn -and $c.UserPrincipalName -ne $AdminUpn))) {
        Write-Host "Exchange Online is signed in as $($c.UserPrincipalName) to $($c.TenantID). Disconnecting." -ForegroundColor Yellow
        Disconnect-ExchangeOnline -Confirm:$false
        $c = $null
    }
    if (-not $c) {
        $p = @{ ShowBanner = $false }; if ($AdminUpn) { $p.UserPrincipalName = $AdminUpn }
        Connect-ExchangeOnline @p
        $c = Get-ConnectionInformation | Where-Object State -eq 'Connected' | Select-Object -First 1
    }
    if ($c.TenantID -ne $TenantId) { throw "Exchange Online is signed in to $($c.TenantID), not $TenantId. Stopping." }
    $c
}
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
    $out = [System.Collections.Generic.List[object]]::new()
    while ($Uri) { $page = Invoke-MgGraphRequest GET $Uri; @($page.value) | ForEach-Object { $out.Add($_) }; $Uri = $page.'@odata.nextLink' }
    , $out.ToArray()
}

# ------------------------------------------------------------ module health (optional helper elsewhere in the repo)
Write-Host ''
Write-Host '== Modules' -ForegroundColor Cyan
$health = Join-Path $PSScriptRoot '..\tools\Test-ModuleHealth.ps1'
if (Test-Path $health) {
    try { & $health | Out-String -Width 250 | Write-Host; Add-Result 'Modules' 'Module health' 'INFO' "Ran $health. Read its output above" }
    catch { Add-Result 'Modules' 'Module health' 'WARN' "Test-ModuleHealth failed: $($_.Exception.Message)" }
}
else { Add-Result 'Modules' 'Module health' 'SKIP' 'tools\Test-ModuleHealth.ps1 not found' }

# ------------------------------------------------------------ snapshot (offline)
Write-Host ''
Write-Host '== Snapshot' -ForegroundColor Cyan
$snapRows = @()
if (-not (Test-Path $SourceSnapshot)) { Add-Result 'Snapshot' 'Snapshot file' 'FAIL' "$SourceSnapshot not found" 'Run Remove-SourceDomain -SnapshotOnly now' }
else {
    $snapRows = @(Import-Csv $SourceSnapshot)
    $withDomain = @($snapRows | Where-Object { "$($_.EmailAddresses)".Split(';') | Where-Object { $_ -match '^smtp:(.+)$' -and (Test-OnDomain $Matches[1]) } }).Count
    $withLeg = @($snapRows | Where-Object LegacyExchangeDN).Count
    if ($withDomain -and $withLeg) { Add-Result 'Snapshot' 'Snapshot is the -pre one' 'PASS' "$($snapRows.Count) rows, $withDomain hold the domain, $withLeg have LegacyExchangeDN" }
    else { Add-Result 'Snapshot' 'Snapshot is the -pre one' 'FAIL' "$withDomain rows hold the domain, $withLeg have LegacyExchangeDN" 'Take a new snapshot before any address is removed (Remove-SourceDomain -SnapshotOnly)' }
}
$mapRows = @()
if (-not (Test-Path $MappingFile)) { Add-Result 'Mapping' 'Mapping file' 'FAIL' "$MappingFile not found" 'Build it with 02b-UserMapping' }
else { $mapRows = @(Import-Csv $MappingFile | Where-Object { $_.Source -and $_.Destination }); Add-Result 'Mapping' 'Mapping file' 'INFO' "$($mapRows.Count) rows" }

# ------------------------------------------------------------ source
Write-Host ''
Write-Host "== Source ($SourceAdmin)" -ForegroundColor Cyan
$exo = Connect-ExoTenant $SourceTenantId $SourceAdmin
$ctx = Connect-GraphTenant $SourceTenantId $SourceAdmin @('User.Read.All', 'Directory.Read.All', 'Application.Read.All', 'Domain.Read.All')
Write-Host "Signed in as $($ctx.Account) to tenant $($ctx.TenantId)" -ForegroundColor Cyan
$srcDomains = Get-GraphAll 'v1.0/domains'
foreach ($d in $Domain) {
    if ($srcDomains.id -contains $d) { Add-Result 'Source' "$d in source" 'INFO' 'Present (it will be released tonight)' }
    else { Add-Result 'Source' "$d in source" 'WARN' 'Not in the source tenant. Already released?' }
}
$deleted = @(Get-GraphAll 'v1.0/directory/deletedItems/microsoft.graph.user?$select=id,userPrincipalName&$top=999' | Where-Object { $_.userPrincipalName -and (Test-OnDomain $_.userPrincipalName) })
if ($deleted) { Add-Result 'Source' 'Deleted users' 'FAIL' "$($deleted.Count) deleted user(s) hold the domain: $(@($deleted.userPrincipalName | Select-Object -First 5) -join ', ')" 'Run Clear-DeletedUserDomain.ps1' }
else { Add-Result 'Source' 'Deleted users' 'PASS' 'None hold the domain' }

$users = Get-GraphAll 'v1.0/users?$select=id,userPrincipalName,mail,proxyAddresses,onPremisesSyncEnabled&$top=999'
$synced = @($users | Where-Object { $_.onPremisesSyncEnabled -and (Test-OnDomain $_.userPrincipalName) })
if ($synced) { Add-Result 'Source' 'Synced users' 'FAIL' "$($synced.Count) synced user(s) on the domain" 'Turn directory sync off for the source first' }
else { Add-Result 'Source' 'Synced users' 'PASS' 'None' }

$recipientIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
Get-EXORecipient -ResultSize Unlimited -Properties ExternalDirectoryObjectId | ForEach-Object { if ($_.ExternalDirectoryObjectId) { [void]$recipientIds.Add("$($_.ExternalDirectoryObjectId)") } }
$proxyOnly = @($users | Where-Object { -not $recipientIds.Contains("$($_.id)") -and (@($_.proxyAddresses | Where-Object { $_ -match '^smtp:(.+)$' -and (Test-OnDomain $Matches[1]) }).Count -or ($_.mail -and (Test-OnDomain $_.mail))) })
if ($proxyOnly) { Add-Result 'Source' 'ProxyOnly users' 'WARN' "$($proxyOnly.Count) user(s) with no mailbox still hold the domain in Entra: $(@($proxyOnly.userPrincipalName | Select-Object -First 5) -join ', ')" 'Graph cannot edit them. Use the admin centre: Settings > Domains > Remove domain wizard' }
else { Add-Result 'Source' 'ProxyOnly users' 'PASS' 'None' }

$apps = Get-GraphAll 'v1.0/applications?$select=id,appId,displayName,identifierUris&$top=999'
$appHits = @($apps | Where-Object { @($_.identifierUris | Where-Object { Test-HostOnDomain $_ }).Count })
if ($appHits) { Add-Result 'Source' 'App identifier URIs' 'FAIL' "$($appHits.Count) app(s): $(@($appHits.displayName | Select-Object -First 5) -join ', ')" 'Remove-SourceDomain -FixApplications, or change the URI to the onmicrosoft.com domain' }
else { Add-Result 'Source' 'App identifier URIs' 'PASS' 'None on the domain' }

# ------------------------------------------------------------ destination
Write-Host ''
Write-Host "== Destination ($DestinationAdmin)" -ForegroundColor Cyan
$exo = Connect-ExoTenant $DestinationTenantId $DestinationAdmin
$ctx = Connect-GraphTenant $DestinationTenantId $DestinationAdmin @('Domain.Read.All', 'Policy.Read.All', 'User.Read.All', 'Directory.Read.All')
Write-Host "Signed in as $($ctx.Account) to tenant $($ctx.TenantId)" -ForegroundColor Cyan

if ($mapRows) {
    $missing = [System.Collections.Generic.List[string]]::new()
    $conflicts = [System.Collections.Generic.List[object]]::new()
    $snapBySource = @{}
    foreach ($s in $snapRows) {
        foreach ($a in "$($s.EmailAddresses)".Split(';')) { if ($a -match '^smtp:(.+)$') { $snapBySource[$Matches[1].ToLower()] = $s } }
    }
    $i = 0
    foreach ($m in $mapRows) {
        $i++; Write-Progress -Activity 'Checking mapping and conflicts' -Status $m.Source -PercentComplete (100 * $i / $mapRows.Count)
        $dest = Get-EXORecipient -Identity $m.Destination.Trim() -Properties ExternalDirectoryObjectId -ErrorAction SilentlyContinue
        if (-not $dest) { $missing.Add($m.Destination); continue }
        $src = $m.Source.Trim().ToLower()
        $addrs = @($src)
        if ($snapBySource.ContainsKey($src)) { $addrs += @("$($snapBySource[$src].EmailAddresses)".Split(';') | Where-Object { $_ -match '^smtp:(.+)$' -and (Test-OnDomain $Matches[1]) } | ForEach-Object { $_.Substring(5).ToLower() }) }
        foreach ($a in @($addrs | Where-Object { Test-OnDomain $_ } | Select-Object -Unique)) {
            $other = @(Get-EXORecipient -Identity $a -Properties ExternalDirectoryObjectId -ErrorAction SilentlyContinue) | Where-Object { "$($_.ExternalDirectoryObjectId)" -ne "$($dest.ExternalDirectoryObjectId)" } | Select-Object -First 1
            if ($other) { $conflicts.Add([pscustomobject]@{ Address = $a; Type = "$($other.RecipientTypeDetails)"; Holder = "$($other.DisplayName) ($($other.ExternalDirectoryObjectId))" }) }
        }
    }
    Write-Progress -Activity 'Checking mapping and conflicts' -Completed
    if ($missing.Count) { Add-Result 'Mapping' 'Destinations resolve' 'FAIL' "$($missing.Count) not found: $(@($missing | Select-Object -First 5) -join ', ')" 'Fix the mapping file or create the recipients' }
    else { Add-Result 'Mapping' 'Destinations resolve' 'PASS' "All $($mapRows.Count) found" }
    $guests = @($conflicts | Where-Object Type -eq 'GuestMailUser')
    $others = @($conflicts | Where-Object Type -ne 'GuestMailUser')
    if ($guests) { Add-Result 'Destination' 'Conflicts with guests' 'WARN' "$($guests.Count) old address(es) held by guests, e.g. $($guests[0].Address) on $($guests[0].Holder)" 'Switch-TargetAddresses -RemoveConflictingGuests, or delete the guests in Entra' }
    if ($others) { Add-Result 'Destination' 'Conflicts with other objects' 'FAIL' "$($others.Count) old address(es) held elsewhere, e.g. $($others[0].Address) on $($others[0].Type) $($others[0].Holder)" 'Decide who should own each address; remove it from the other object' }
    if (-not $conflicts.Count) { Add-Result 'Destination' 'Conflicts' 'PASS' 'No old address is held by another object' }
    $conflicts | ForEach-Object { $results.Add([pscustomobject]@{ Area = 'Conflict'; Check = $_.Address; Result = 'INFO'; Detail = "$($_.Type) $($_.Holder)"; Fix = '' }) }
}

$txtByDomain = @{}
foreach ($d in $Domain) {
    $dom = try { Invoke-MgGraphRequest GET "v1.0/domains/$d" } catch { $null }
    if (-not $dom) { Add-Result 'Destination' "$d state" 'WARN' 'Not added yet' 'Run Add-TargetDomain.ps1 now to get the MS= TXT, and publish it before the night'; continue }
    if ($dom.isVerified) { Add-Result 'Destination' "$d state" 'PASS' 'Verified'; continue }
    $txt = ((Invoke-MgGraphRequest GET "v1.0/domains/$d/verificationDnsRecords").value | Where-Object { $_.recordType -eq 'Txt' } | Select-Object -First 1).text
    $txtByDomain[$d] = $txt
    Add-Result 'Destination' "$d state" 'PASS' "Added, unverified. TXT to publish: $txt"
}

$ca = @(try { Get-GraphAll 'v1.0/identity/conditionalAccess/policies' } catch { Add-Result 'Destination' 'Conditional Access' 'SKIP' "Couldn't read policies: $($_.Exception.Message)" })
$locCa = @($ca | Where-Object { $_.state -ne 'disabled' -and ((@($_.conditions.locations.includeLocations).Count) -or (@($_.conditions.locations.excludeLocations).Count)) })
if ($locCa) { Add-Result 'Destination' 'Location-based Conditional Access' 'WARN' "$($locCa.Count) policy(ies): $(@($locCa.displayName | Select-Object -First 4) -join '; ')" 'Your test sign-in from your real location may be blocked. Know which named locations apply before the night' }
elseif ($ca) { Add-Result 'Destination' 'Location-based Conditional Access' 'PASS' 'No enabled policy uses locations' }

# ------------------------------------------------------------ DNS
Write-Host ''
Write-Host '== DNS' -ForegroundColor Cyan
if (-not (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue)) { Add-Result 'DNS' 'DNS checks' 'SKIP' 'Resolve-DnsName not available (Windows only)' }
else {
    foreach ($d in $Domain) {
        $ns = @(Resolve-DnsName -Name $d -Type NS -DnsOnly -ErrorAction SilentlyContinue | Where-Object { $_.Section -eq 'Answer' -and "$($_.Type)" -eq 'NS' })
        if (-not $ns) { Add-Result 'DNS' "$d name servers" 'FAIL' 'No NS records found'; continue }
        $hosts = @($ns | ForEach-Object { "$($_.NameHost)".TrimEnd('.') } | Sort-Object -Unique)
        Add-Result 'DNS' "$d name servers" 'INFO' "DNS is hosted on $($hosts -join ', '). Make sure you can log in there"
        $ttl = ($ns | Measure-Object TTL -Maximum).Maximum
        if ($ttl -ge 86400) { Add-Result 'DNS' "$d NS TTL" 'WARN' "$ttl s. Moving DNS hosting now would not settle before cutover" 'If you need to move DNS, do it at least that long before' }
        else { Add-Result 'DNS' "$d NS TTL" 'PASS' "$ttl s" }
        if ($txtByDomain[$d]) {
            $missingOn = @($hosts | Where-Object { $h = $_; -not (@(Resolve-DnsName -Name $d -Type TXT -Server $h -DnsOnly -ErrorAction SilentlyContinue | Where-Object { $_.Section -eq 'Answer' } | ForEach-Object { $_.Strings -join '' }) -contains $txtByDomain[$d]) })
            if ($missingOn) { Add-Result 'DNS' "$d MS= TXT published" 'WARN' "Not yet on $($missingOn -join ', ')" 'Publish it now: Add-TargetDomain -WaitForDns tells you when every server has it' }
            else { Add-Result 'DNS' "$d MS= TXT published" 'PASS' "On all $($hosts.Count) name servers" }
        }
    }
}

# ------------------------------------------------------------ summary
$file = ".\CutoverReadiness-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
$results | Export-Csv $file -NoTypeInformation -WhatIf:$false
Write-Host ''
$results | Where-Object Area -ne 'Conflict' | Format-Table Area, Check, Result, Detail, Fix -AutoSize -Wrap | Out-String -Width 250 | Write-Host
$fails = @($results | Where-Object Result -eq 'FAIL').Count
$warns = @($results | Where-Object Result -eq 'WARN').Count
if ($fails) { Write-Host "NO GO: $fails FAIL, $warns WARN. Report: $file" -ForegroundColor Red }
else { Write-Host "GO: no FAIL ($warns WARN to read). Report: $file" -ForegroundColor Green }
