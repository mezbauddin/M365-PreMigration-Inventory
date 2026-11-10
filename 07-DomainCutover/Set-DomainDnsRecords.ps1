<#
.SYNOPSIS
    Writes DNS records to a GoDaddy or Cloudflare zone through the provider's API, one Type+Name group at a time,
    after backing up the whole zone to JSON.

.DESCRIPTION
    TL;DR: Give it a CSV of the records you want (Type, Name, Value, TTL, Priority), or the CSV from
    Export-DomainDns.ps1 -FullScan, and it makes each Type+Name group in the zone match exactly. It backs up the zone
    first, never touches NS or SOA, refuses any apex CNAME, and leaves alone anything in the zone that isn't in your
    file (it lists those at the end). -WhatIf prints the planned groups and changes nothing.

    When to use it:
      - Replicating a zone you can't log in to (found with Export-DomainDns -FullScan) into one you can, before you
        move the name servers. Do that at least one NS TTL before cutover (often 48 hours).
      - Applying the Change and Add rows from Add-TargetDomain on the night, without clicking through a portal at
        speed.

    How it writes:
      GoDaddy     GET /v1/domains/{domain}/records for the backup, then PUT /v1/domains/{domain}/records/{type}/{name}
                  with every record for that group. NS and SOA are never sent, and the PUT replaces GoDaddy's default
                  "Parked" A record at @ cleanly.
      Cloudflare  Looks up the zone ID, lists records, then per group deletes records that aren't in your set and
                  creates the missing ones. New records are DNS only (not proxied).

    Credentials are never stored. GoDaddy: env GODADDY_KEY and GODADDY_SECRET, or it prompts (the secret hidden).
    Cloudflare: env CLOUDFLARE_API_TOKEN, or it prompts. Use a token scoped to Zone.DNS edit on that zone only.

    Trap: a GoDaddy API key works on every domain in the account. Delete it in the developer portal when you're done.
    Trap: GoDaddy adds its own _domainconnect CNAME. Leave it; it's listed as "not in file" and nothing more.
    Trap: an apex CNAME (for example @ -> <tenant>.sharepoint.com, which the Microsoft 365 wizard suggests) would
    break MX and TXT at the root. The script refuses it.
    Tip: GoDaddy's minimum TTL is 600 seconds; lower values are raised to 600.
    Trap: GoDaddy keeps SRV records under separate service and protocol fields, so for SRV the script only adds
    missing records (PATCH). Delete stale SRV values in the portal. Teams doesn't need the old Skype SRV records.
    No script: the provider's DNS page. Fine for five records; risky for fifty at 9pm.

.PARAMETER Domain
    The zone, e.g. fabrikam.com.

.PARAMETER Provider
    GoDaddy or Cloudflare.

.PARAMETER RecordsFile
    CSV with Type, Name, Value, TTL, Priority. Name is relative (@, www, selector1._domainkey) or a full name. MX
    host goes in Value and preference in Priority. SRV Value is "priority weight port target" (or "weight port
    target" with Priority filled).

.PARAMETER FromScan
    CSV from Export-DomainDns.ps1 -FullScan (Domain, Name, Type, TTL, Value). Rows for other domains, NS and SOA are
    skipped. MX Value is "preference host".

.PARAMETER BackupFolder
    Where the zone backup JSON goes. Default the current folder.

.EXAMPLE
    .\Set-DomainDnsRecords.ps1 -Domain fabrikam.com -Provider GoDaddy -FromScan .\DomainDns-FullScan-20261108-1000.csv -WhatIf

.EXAMPLE
    $env:CLOUDFLARE_API_TOKEN = Read-Host 'Token'
    .\Set-DomainDnsRecords.ps1 -Domain fabrikam.com -Provider Cloudflare -RecordsFile .\cutover-records.csv

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.0.0
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: PowerShell 7 or Windows PowerShell 5.1. A GoDaddy production API key and secret, or a Cloudflare API
              token with DNS edit on the zone.
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'File')]
param(
    [Parameter(Mandatory)][string]$Domain,
    [Parameter(Mandatory)][ValidateSet('GoDaddy', 'Cloudflare')][string]$Provider,
    [Parameter(Mandatory, ParameterSetName = 'File')][string]$RecordsFile,
    [Parameter(Mandatory, ParameterSetName = 'Scan')][string]$FromScan,
    [string]$BackupFolder = '.'
)

$ErrorActionPreference = 'Stop'
$Domain = $Domain.Trim().ToLower().TrimEnd('.')

# ------------------------------------------------------------ read and normalise the wanted records
function ConvertTo-RelativeName([string]$Name) {
    $n = $Name.Trim().ToLower().TrimEnd('.')
    if (-not $n -or $n -eq '@' -or $n -eq $Domain) { return '@' }
    if ($n.EndsWith(".$Domain")) { return $n.Substring(0, $n.Length - $Domain.Length - 1) }
    $n
}
function Get-WantedRecords {
    $src = if ($PSCmdlet.ParameterSetName -eq 'Scan') { $FromScan } else { $RecordsFile }
    if (-not (Test-Path $src)) { throw "$src not found." }
    foreach ($row in Import-Csv $src) {
        if ($PSCmdlet.ParameterSetName -eq 'Scan' -and $row.Domain -and $row.Domain.Trim().ToLower() -ne $Domain) { continue }
        $type = "$($row.Type)".Trim().ToUpper()
        if (-not $type -or $type -in 'NS', 'SOA') { continue }
        $name = ConvertTo-RelativeName "$($row.Name)"
        $value = "$($row.Value)".Trim()
        if (-not $value -or $value -eq '(none)') { continue }
        $ttl = if ("$($row.TTL)" -match '^\d+$') { [int]$row.TTL } else { 3600 }
        $prio = if ("$($row.Priority)" -match '^\d+$') { [int]$row.Priority } else { $null }
        $weight = $null; $port = $null
        if ($type -eq 'MX' -and $value -match '^(\d+)\s+(\S+)$') { $prio = [int]$Matches[1]; $value = $Matches[2] }
        if ($type -eq 'SRV') {
            $parts = $value -split '\s+'
            if ($parts.Count -eq 4) { $prio = [int]$parts[0]; $weight = [int]$parts[1]; $port = [int]$parts[2]; $value = $parts[3] }
            elseif ($parts.Count -eq 3) { $weight = [int]$parts[0]; $port = [int]$parts[1]; $value = $parts[2] }
            else { throw "SRV value '$($row.Value)' for $name should be 'priority weight port target'." }
            if ($null -eq $prio) { $prio = 0 }
        }
        if ($type -in 'CNAME', 'MX', 'SRV', 'NS') { $value = $value.TrimEnd('.').ToLower() }
        if ($type -eq 'TXT') { $value = $value.Trim('"') }
        [pscustomobject]@{ Type = $type; Name = $name; Value = $value; TTL = $ttl; Priority = $prio; Weight = $weight; Port = $port }
    }
}
$wanted = @(Get-WantedRecords)
if (-not $wanted) { throw 'No records to write.' }
$apexCname = @($wanted | Where-Object { $_.Type -eq 'CNAME' -and $_.Name -eq '@' })
if ($apexCname) { throw "Refusing: the file has an apex CNAME (@ -> $($apexCname[0].Value)). That would break MX and TXT for the whole domain. Remove it." }
foreach ($g in $wanted | Group-Object Name) {
    if (($g.Group.Type -contains 'CNAME') -and @($g.Group | Where-Object Type -ne 'CNAME').Count) { throw "Refusing: $($g.Name) has a CNAME and other records. A CNAME can't share a name." }
}
$fastApex = @($wanted | Where-Object { $_.Type -eq 'A' -and $_.Name -eq '@' -and $_.TTL -le 60 })
if ($fastApex) {
    Write-Warning "Apex A record(s) with TTL <= 60 s ($(@($fastApex.Value) -join ', ')). That usually means the old host used an alias record (for example to a CDN). Static copies work today and break when the CDN changes address. Ask the website owner for the right target."
}
$groups = @($wanted | Group-Object Type, Name)
Write-Host "$($wanted.Count) records in $($groups.Count) Type+Name groups for $Domain on $Provider." -ForegroundColor Cyan

function Get-Key($Type, $Value, $Prio, $Weight, $Port) {
    $v = "$Value".Trim().TrimEnd('.').Trim('"')
    if ($Type -ne 'TXT') { $v = $v.ToLower() }
    switch ($Type) {
        'MX'    { "$Prio $v" }
        'SRV'   { "$Prio $Weight $Port $v" }
        default { $v }
    }
}
function Get-Secret([string]$EnvName, [string]$Prompt, [switch]$Hidden) {
    $v = [Environment]::GetEnvironmentVariable($EnvName)
    if ($v) { return $v }
    if ($Hidden) { $s = Read-Host $Prompt -AsSecureString; return [System.Net.NetworkCredential]::new('', $s).Password }
    Read-Host $Prompt
}

# ------------------------------------------------------------ provider plumbing
if ($Provider -eq 'GoDaddy') {
    $key = Get-Secret 'GODADDY_KEY' 'GoDaddy API key'
    $secret = Get-Secret 'GODADDY_SECRET' 'GoDaddy API secret' -Hidden
    $headers = @{ Authorization = "sso-key $($key):$($secret)"; Accept = 'application/json' }
    $base = "https://api.godaddy.com/v1/domains/$Domain/records"
    function Get-Zone {
        $resp = Invoke-RestMethod -Method Get -Uri $base -Headers $headers
        foreach ($r in @($resp | ForEach-Object { $_ })) {
            $name = "$($r.name)".ToLower()
            # GoDaddy keeps SRV service and protocol in their own fields
            if ($r.type -eq 'SRV' -and $r.service -and -not $name.StartsWith("$($r.service).".ToLower())) {
                $name = (@($r.service, $r.protocol) + $(if ($name -and $name -ne '@') { $name })) -join '.'
            }
            [pscustomobject]@{ Type = "$($r.type)"; Name = $name; Value = "$($r.data)"; TTL = $r.ttl; Priority = $r.priority; Weight = $r.weight; Port = $r.port; Raw = $r }
        }
    }
    function Set-Group($Type, $Name, $Records) {
        $body = foreach ($r in $Records) {
            $ttl = [math]::Max(600, [int]$r.TTL)
            $o = [ordered]@{ data = $r.Value; ttl = $ttl }
            if ($Type -eq 'MX') { $o.priority = [int]$r.Priority }
            if ($Type -eq 'SRV') {
                $bits = $Name -split '\.', 3
                $o.service = $bits[0]; $o.protocol = $bits[1]; $o.name = $(if ($bits.Count -gt 2) { $bits[2] } else { '@' })
                $o.priority = [int]$r.Priority; $o.weight = [int]$r.Weight; $o.port = [int]$r.Port
            }
            $o
        }
        if ($Type -eq 'SRV') {
            # GoDaddy files SRV records under service/protocol fields, not the name in the URL, so add missing ones
            # with PATCH instead of replacing the group. Old SRV values are left for you to delete in the portal.
            $haveKeys = @($script:zoneNow | Where-Object { $_.Type -eq 'SRV' -and $_.Name -eq $Name } | ForEach-Object { Get-Key 'SRV' $_.Value $_.Priority $_.Weight $_.Port })
            $missing = @(for ($i = 0; $i -lt @($Records).Count; $i++) {
                    $r = @($Records)[$i]
                    if ($haveKeys -notcontains (Get-Key 'SRV' $r.Value $r.Priority $r.Weight $r.Port)) { $o = @($body)[$i]; $o.type = 'SRV'; $o } })
            if ($missing) { Invoke-RestMethod -Method Patch -Uri $base -Headers $headers -ContentType 'application/json' -Body (ConvertTo-Json -InputObject $missing -Depth 4 -Compress) | Out-Null }
            return
        }
        $json = ConvertTo-Json -InputObject @($body) -Depth 4 -Compress
        Invoke-RestMethod -Method Put -Uri "$base/$Type/$([uri]::EscapeDataString($Name))" -Headers $headers -ContentType 'application/json' -Body $json | Out-Null
    }
}
else {
    $token = Get-Secret 'CLOUDFLARE_API_TOKEN' 'Cloudflare API token' -Hidden
    $headers = @{ Authorization = "Bearer $token" }
    $api = 'https://api.cloudflare.com/client/v4'
    $zone = (Invoke-RestMethod -Method Get -Uri "$api/zones?name=$Domain" -Headers $headers).result | Select-Object -First 1
    if (-not $zone) { throw "Cloudflare zone $Domain not found for this token." }
    $zoneId = $zone.id
    function Get-Zone {
        $page = 1
        do {
            $resp = Invoke-RestMethod -Method Get -Uri "$api/zones/$zoneId/dns_records?per_page=100&page=$page" -Headers $headers
            foreach ($r in $resp.result) {
                $v = "$($r.content)"; $w = $null; $pt = $null; $pr = $r.priority
                if ($r.type -eq 'SRV' -and $r.data) { $v = "$($r.data.target)"; $w = $r.data.weight; $pt = $r.data.port; $pr = $r.data.priority }
                [pscustomobject]@{ Type = "$($r.type)"; Name = (ConvertTo-RelativeName "$($r.name)"); Value = $v; TTL = $r.ttl; Priority = $pr; Weight = $w; Port = $pt; Raw = $r }
            }
            $page++
        } while ($resp.result_info -and $resp.result_info.page -lt $resp.result_info.total_pages)
    }
    function Set-Group($Type, $Name, $Records) {
        $fqdn = if ($Name -eq '@') { $Domain } else { "$Name.$Domain" }
        $current = @($script:zoneNow | Where-Object { $_.Type -eq $Type -and $_.Name -eq $Name })
        $wantKeys = @($Records | ForEach-Object { Get-Key $Type $_.Value $_.Priority $_.Weight $_.Port })
        foreach ($c in $current) {
            if ($wantKeys -notcontains (Get-Key $Type $c.Value $c.Priority $c.Weight $c.Port)) {
                Invoke-RestMethod -Method Delete -Uri "$api/zones/$zoneId/dns_records/$($c.Raw.id)" -Headers $headers | Out-Null
            }
        }
        $haveKeys = @($current | ForEach-Object { Get-Key $Type $_.Value $_.Priority $_.Weight $_.Port })
        foreach ($r in $Records) {
            if ($haveKeys -contains (Get-Key $Type $r.Value $r.Priority $r.Weight $r.Port)) { continue }
            $o = [ordered]@{ type = $Type; name = $fqdn; ttl = [math]::Max(60, [int]$r.TTL); proxied = $false }
            if ($Type -eq 'SRV') { $o.data = @{ priority = [int]$r.Priority; weight = [int]$r.Weight; port = [int]$r.Port; target = $r.Value } }
            else { $o.content = $r.Value }
            if ($Type -eq 'MX') { $o.priority = [int]$r.Priority }
            if ($Type -notin 'A', 'AAAA', 'CNAME') { $o.Remove('proxied') }
            Invoke-RestMethod -Method Post -Uri "$api/zones/$zoneId/dns_records" -Headers $headers -ContentType 'application/json' -Body ($o | ConvertTo-Json -Depth 4 -Compress) | Out-Null
        }
    }
}

# ------------------------------------------------------------ back up, plan, apply
$script:zoneNow = @(Get-Zone)
if (-not (Test-Path $BackupFolder)) { New-Item -ItemType Directory -Path $BackupFolder -WhatIf:$false | Out-Null }
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = Join-Path $BackupFolder "DnsBackup-$Domain-$Provider-$stamp.json"
ConvertTo-Json -InputObject @($script:zoneNow | ForEach-Object { $_.Raw }) -Depth 6 | Set-Content -Path $backup -Encoding UTF8 -WhatIf:$false
Write-Host "Backed up $($script:zoneNow.Count) records to $backup" -ForegroundColor Green

$log = [System.Collections.Generic.List[object]]::new()
foreach ($g in $groups) {
    $type = $g.Group[0].Type; $name = $g.Group[0].Name
    $want = @($g.Group)
    $have = @($script:zoneNow | Where-Object { $_.Type -eq $type -and $_.Name -eq $name })
    $wantKeys = @($want | ForEach-Object { Get-Key $type $_.Value $_.Priority $_.Weight $_.Port } | Sort-Object -Unique)
    $haveKeys = @($have | ForEach-Object { Get-Key $type $_.Value $_.Priority $_.Weight $_.Port } | Sort-Object -Unique)
    $label = "$type $name"
    if (-not (Compare-Object $wantKeys $haveKeys) -and $have.Count) {
        $log.Add([pscustomobject]@{ Group = $label; Status = 'Unchanged'; Records = ($wantKeys -join ' | '); Was = '' }); continue
    }
    $was = $haveKeys -join ' | '
    if ($PSCmdlet.ShouldProcess("$Domain $label", "Set to: $($wantKeys -join ' | ')")) {
        try { Set-Group $type $name $want; $log.Add([pscustomobject]@{ Group = $label; Status = 'Set'; Records = ($wantKeys -join ' | '); Was = $was }) }
        catch { $log.Add([pscustomobject]@{ Group = $label; Status = 'Problem'; Records = ($wantKeys -join ' | '); Was = "$was ($($_.Exception.Message))" }) }
    }
    else { $log.Add([pscustomobject]@{ Group = $label; Status = 'WhatIf'; Records = ($wantKeys -join ' | '); Was = $was }) }
}
$log | Format-Table Group, Status, Records, Was -AutoSize -Wrap | Out-String -Width 250 | Write-Host

# ------------------------------------------------------------ final zone and leftovers
if (-not $WhatIfPreference) { $script:zoneNow = @(Get-Zone) }
$wantedGroups = @($groups | ForEach-Object { "$($_.Group[0].Type)|$($_.Group[0].Name)" })
$extra = @($script:zoneNow | Where-Object { $_.Type -notin 'NS', 'SOA' -and $wantedGroups -notcontains "$($_.Type)|$($_.Name)" })
Write-Host "$(if ($WhatIfPreference) { 'Zone as it stands (-WhatIf, nothing changed)' } else { 'Zone now' }) ($($script:zoneNow.Count) records):" -ForegroundColor Cyan
$script:zoneNow | Sort-Object Name, Type | Format-Table Type, Name, Value, TTL, Priority -AutoSize | Out-String -Width 250 | Write-Host
if ($extra) {
    Write-Host 'In the zone but not in your file (left alone):' -ForegroundColor Yellow
    $extra | Format-Table Type, Name, Value -AutoSize | Out-String -Width 250 | Write-Host
}
$file = Join-Path $BackupFolder "SetDomainDnsRecords-$Domain-$stamp.csv"
$log | Export-Csv $file -NoTypeInformation -WhatIf:$false
Write-Host "Report: $file   Backup: $backup" -ForegroundColor Cyan
if ($Provider -eq 'GoDaddy') { Write-Host 'Trap: the GoDaddy API key works on every domain in the account. Delete it when you are done.' -ForegroundColor Yellow }
