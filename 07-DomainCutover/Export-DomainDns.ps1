<#
.SYNOPSIS
    Records the current DNS for the domains you're moving, so you know what to change at cutover and what to put
    back if you have to roll back. With -FullScan, asks the authoritative name server for the whole zone it can find.

.DESCRIPTION
    TL;DR: Read-only. By default it looks up the records Microsoft 365 depends on (MX, SPF and other TXT,
    autodiscover, DKIM selectors, DMARC, device registration, Teams and Skype records) and writes them, with their
    TTLs, to a CSV. Run it a few days before cutover to see the TTLs to lower, and again on the night as your
    rollback copy.

    -FullScan is for when you can't log in to wherever DNS is hosted, or you plan to move DNS somewhere you can. It
    finds the authoritative name servers, asks one of them directly for the apex A, AAAA, MX, TXT, NS and SOA
    records, then tries about 100 common names (or your own -Wordlist) for CNAME, A, AAAA, TXT, MX and SRV. The CSV
    (Domain, Name, Type, TTL, Value) feeds straight into Set-DomainDnsRecords.ps1 -FromScan. It also tells you:
      - where DNS is hosted (the NS records). That is often not the registrar. On a live cutover it was a cloud DNS
        service nobody had the login for, while everyone was looking at the registrar.
      - the NS TTL. 172800 (48 hours) is common: a name server change needs that long to settle everywhere.
      - apex A records with a TTL of 60 seconds or less. That pattern usually means an alias record (for example a
        cloud DNS alias to a CDN). Copying those addresses as static A records works today and breaks when the CDN
        moves the site.

    Trap: a wordlist scan only finds names it asks for. Zone transfers are almost always refused, so anything with an
    unusual name (a vendor's verification record, a one-off CNAME) is missed. Ask whoever runs the website and any
    SaaS that sends mail as the domain, and add their names to -Wordlist.
    Tip: CAA records aren't scanned (Windows Resolve-DnsName doesn't support them). Check them in a web DNS lookup.
    Paid tool: DNS hosting providers and security scanners can export or enumerate a zone; use them if you have them.
    No script: nslookup -type=any against the authoritative server, one name at a time.

.PARAMETER Domain
    The domains moving between tenants.

.PARAMETER Server
    Optional DNS server to ask, e.g. 1.1.1.1 or 8.8.8.8, to see what the internet sees rather than a local cache.
    Used for the default check and to find the name servers in -FullScan.

.PARAMETER MaxTtl
    TTL in seconds above which a record is flagged for lowering before cutover. Default 3600.

.PARAMETER FullScan
    Ask the authoritative name server directly for the apex records and every name in the wordlist.

.PARAMETER Wordlist
    Optional text file, one name per line (relative to the domain, e.g. www or selector1._domainkey), used with
    -FullScan in addition to the built-in list.

.EXAMPLE
    .\Export-DomainDns.ps1 -Domain fabrikam.com, fabrikam.co.uk -Server 1.1.1.1

.EXAMPLE
    .\Export-DomainDns.ps1 -Domain fabrikam.com -FullScan -Wordlist .\extra-names.txt

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.1.0
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: Windows (Resolve-DnsName). Changes nothing.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string[]]$Domain,
    [string]$Server,
    [int]$MaxTtl = 3600,
    [switch]$FullScan,
    [string]$Wordlist
)

$ErrorActionPreference = 'Stop'
if (-not (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue)) { throw 'Resolve-DnsName is needed (Windows, DnsClient module).' }

function Format-DnsValue($r) {
    # $r is one Resolve-DnsName record. Inside a switch $_ is the switch value, so never use $_ for the record here.
    switch ("$($r.Type)") {
        'MX'    { "$($r.Preference) $("$($r.NameExchange)".TrimEnd('.'))" }
        'TXT'   { ($r.Strings -join '') }
        'CNAME' { "$($r.NameHost)".TrimEnd('.') }
        'NS'    { "$($r.NameHost)".TrimEnd('.') }
        'SRV'   { "$($r.Priority) $($r.Weight) $($r.Port) $("$($r.NameTarget)".TrimEnd('.'))" }
        'SOA'   { "$("$($r.PrimaryServer)".TrimEnd('.')) $("$($r.NameAdministrator)".TrimEnd('.')) $($r.SerialNumber)" }
        'A'     { "$($r.IPAddress)" }
        'AAAA'  { "$($r.IPAddress)" }
        default { "$($r.Type)" }
    }
}
function Get-Answer([string]$Name, [string]$Type, [string]$Srv) {
    $p = @{ Name = $Name; Type = $Type; DnsOnly = $true; QuickTimeout = $true; ErrorAction = 'SilentlyContinue' }
    if ($Srv) { $p.Server = $Srv }
    @(Resolve-DnsName @p | Where-Object { $_.Section -eq 'Answer' -and "$($_.Type)" -eq $Type -and "$($_.Name)".TrimEnd('.') -eq $Name })
}

$stamp = Get-Date -Format 'yyyyMMdd-HHmm'

# ================================================================== full scan
if ($FullScan) {
    $names = @(
        'www', 'mail', 'webmail', 'smtp', 'pop', 'pop3', 'imap', 'mx', 'mx1', 'mx2', 'ftp', 'sftp', 'autodiscover',
        'enterpriseregistration', 'enterpriseenrollment', 'lyncdiscover', 'sip', 'msoid', 'meet', 'dialin', 'teams',
        '_dmarc', 'selector1._domainkey', 'selector2._domainkey', 'google._domainkey', 'k1._domainkey', 'k2._domainkey',
        's1._domainkey', 's2._domainkey', 'default._domainkey', 'dkim._domainkey', 'mail._domainkey', 'smtp._domainkey',
        'mandrill._domainkey', 'mxvault._domainkey', 'em', 'email', 'bounce', 'bounces', 'links', 'click', 'track',
        'news', 'newsletter', 'marketing', 'portal', 'vpn', 'remote', 'rdp', 'gateway', 'owa', 'exchange', 'cpanel',
        'whm', 'webdisk', 'admin', 'blog', 'shop', 'store', 'app', 'apps', 'api', 'dev', 'test', 'staging', 'beta',
        'demo', 'docs', 'help', 'support', 'status', 'login', 'sso', 'auth', 'id', 'secure', 'files', 'cdn', 'static',
        'assets', 'img', 'images', 'media', 'video', 'm', 'mobile', 'crm', 'erp', 'hr', 'intranet', 'extranet', 'wiki',
        'calendar', 'cal', 'office', 'host', 'server', 'ns1', 'ns2', 'dns', 'cloud', 'go', 'my', 'pay', 'billing',
        'forms', 'survey', 'booking', 'jobs', 'careers', 'old', 'new', 'legacy', 'backup', 'zoom', 'citrix',
        '_sip._tls', '_sipfederationtls._tcp', '_autodiscover._tcp', '_caldav._tcp', '_carddav._tcp', '_submission._tcp',
        '_imaps._tcp', '_xmpp-client._tcp', '_domainconnect'
    )
    if ($Wordlist) {
        if (-not (Test-Path $Wordlist)) { throw "Wordlist $Wordlist not found." }
        $names += @(Get-Content $Wordlist | ForEach-Object { $_.Trim().ToLower() } | Where-Object { $_ -and $_ -notlike '#*' })
    }
    $names = @($names | Select-Object -Unique)

    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($d in $Domain) {
        $d = $d.ToLower().Trim()
        Write-Host ''
        Write-Host "== $d" -ForegroundColor Cyan
        $nsRecords = @(Get-Answer $d 'NS' $Server)
        if (-not $nsRecords) { Write-Host "  No NS records found for $d. Is it a registered zone?" -ForegroundColor Red; continue }
        $nsHosts = @($nsRecords | ForEach-Object { "$($_.NameHost)".TrimEnd('.') } | Sort-Object -Unique)
        Write-Host "  DNS is hosted on: $($nsHosts -join ', ')" -ForegroundColor Yellow
        $nsTtl = ($nsRecords | Measure-Object TTL -Maximum).Maximum
        if ($nsTtl -ge 86400) {
            Write-Host "  NS TTL is $nsTtl s ($([math]::Round($nsTtl / 3600)) h). A name server move needs that long to settle. Move DNS at least that long before cutover." -ForegroundColor Yellow
        }
        # Pick the first authoritative server that answers
        $auth = $null
        foreach ($h in $nsHosts) { if (Get-Answer $d 'SOA' $h) { $auth = $h; break } }
        if (-not $auth) { Write-Host '  None of the name servers answered directly. Using the normal resolver; TTLs may be cached values.' -ForegroundColor Yellow }
        else { Write-Host "  Asking $auth directly." -ForegroundColor Gray }

        $add = {
            param($name, $answers)
            foreach ($r in $answers) {
                $rows.Add([pscustomobject]@{ Domain = $d; Name = $name; Type = "$($r.Type)"; TTL = $r.TTL; Value = (Format-DnsValue $r) })
            }
        }
        foreach ($t in 'SOA', 'NS', 'A', 'AAAA', 'MX', 'TXT') { & $add $d (Get-Answer $d $t $auth) }

        $i = 0
        foreach ($n in $names) {
            $i++
            Write-Progress -Activity "Scanning $d" -Status $n -PercentComplete (100 * $i / $names.Count)
            $fqdn = "$n.$d"
            $cname = @(Get-Answer $fqdn 'CNAME' $auth)
            if ($cname) { & $add $fqdn $cname; continue }   # a CNAME can't sit with other records
            $types = if ($n.StartsWith('_') -and $n -match '\._(tcp|udp|tls)$') { @('SRV') } else { @('A', 'AAAA', 'TXT', 'MX') }
            foreach ($t in $types) { & $add $fqdn (Get-Answer $fqdn $t $auth) }
        }
        Write-Progress -Activity "Scanning $d" -Completed

        $fast = @($rows | Where-Object { $_.Domain -eq $d -and $_.Name -eq $d -and $_.Type -eq 'A' -and [int]$_.TTL -le 60 })
        if ($fast) {
            Write-Host "  Apex A record(s) with TTL <= 60 s ($(@($fast.Value) -join ', ')): looks like an alias record, e.g. a cloud DNS alias to a CDN. Copying as static A records is a risk." -ForegroundColor Yellow
        }
        Write-Host "  $(@($rows | Where-Object Domain -eq $d).Count) records found." -ForegroundColor Green
    }

    $file = ".\DomainDns-FullScan-$stamp.csv"
    $rows | Select-Object Domain, Name, Type, TTL, Value | Export-Csv $file -NoTypeInformation -WhatIf:$false
    $rows | Format-Table Domain, Name, Type, TTL, Value -AutoSize -Wrap | Out-String -Width 250 | Write-Host
    Write-Host "Saved to $file. A wordlist scan only finds names it asks for: check with the website and mail-sending owners." -ForegroundColor Cyan
    return
}

# ================================================================== Microsoft 365 records (default)
# Name relative to the domain, record type, why it matters
$checks = @(
    @('@', 'MX', 'Inbound mail'),
    @('@', 'TXT', 'SPF and domain verification'),
    @('autodiscover', 'CNAME', 'Outlook setup'),
    @('selector1._domainkey', 'CNAME', 'DKIM signing'),
    @('selector2._domainkey', 'CNAME', 'DKIM signing'),
    @('_dmarc', 'TXT', 'DMARC policy'),
    @('enterpriseregistration', 'CNAME', 'Device registration'),
    @('enterpriseenrollment', 'CNAME', 'Intune enrolment'),
    @('lyncdiscover', 'CNAME', 'Teams and Skype (legacy)'),
    @('sip', 'CNAME', 'Teams and Skype (legacy)'),
    @('_sip._tls', 'SRV', 'Teams and Skype (legacy)'),
    @('_sipfederationtls._tcp', 'SRV', 'Teams federation (legacy)')
)

$rows = foreach ($d in $Domain) {
    $d = $d.ToLower().Trim()
    foreach ($c in $checks) {
        $name = if ($c[0] -eq '@') { $d } else { "$($c[0]).$d" }
        $answers = Get-Answer $name $c[1] $Server
        if (-not $answers) {
            [pscustomobject]@{ Domain = $d; Name = $name; Type = $c[1]; Value = '(none)'; TTL = ''; Purpose = $c[2]; Flag = '' }
            continue
        }
        foreach ($a in $answers) {
            $flag = if ($a.TTL -gt $MaxTtl) { "Lower TTL (currently $($a.TTL)s)" } else { '' }
            [pscustomobject]@{ Domain = $d; Name = $name; Type = $c[1]; Value = (Format-DnsValue $a); TTL = $a.TTL; Purpose = $c[2]; Flag = $flag }
        }
    }
}

$file = ".\DomainDns-$stamp.csv"
$rows | Export-Csv $file -NoTypeInformation -WhatIf:$false
$rows | Format-Table Domain, Type, Name, Value, TTL, Flag -AutoSize -Wrap | Out-String -Width 250 | Write-Host
$high = @($rows | Where-Object Flag)
if ($high) { Write-Host "$($high.Count) record(s) have a TTL above $MaxTtl seconds. Lower them at least one old-TTL period before cutover." -ForegroundColor Yellow }
Write-Host "Saved to $file. Keep it: it's your rollback reference. For the whole zone, run again with -FullScan." -ForegroundColor Cyan
