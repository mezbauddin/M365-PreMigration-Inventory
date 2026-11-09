<#
.SYNOPSIS
    Records the current public DNS for the domains you're moving, so you know what to change at cutover and what to
    put back if you have to roll back.

.DESCRIPTION
    TL;DR: Read-only. Looks up the records Microsoft 365 depends on for each domain (MX, SPF and other TXT, autodiscover,
    DKIM selectors, DMARC, device registration, Teams and Skype records) and writes them, with their TTLs, to a CSV. Run
    it a few days before cutover to see the TTLs you need to lower, and again on the night as your rollback copy.

.PARAMETER Domain
    The domains moving between tenants.

.PARAMETER Server
    Optional DNS server to ask, e.g. 1.1.1.1 or 8.8.8.8, to see what the internet sees rather than a local cache.

.PARAMETER MaxTtl
    TTL in seconds above which a record is flagged for lowering before cutover. Default 3600.

.EXAMPLE
    .\Export-DomainDns.ps1 -Domain fabrikam.com, fabrikam.co.uk -Server 1.1.1.1

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.0.0
    Date    : 2026-11-09
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: Windows (Resolve-DnsName). Changes nothing.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string[]]$Domain,
    [string]$Server,
    [int]$MaxTtl = 3600
)

$ErrorActionPreference = 'Stop'
if (-not (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue)) { throw 'Resolve-DnsName is needed (Windows, DnsClient module).' }

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
        $p = @{ Name = $name; Type = $c[1]; DnsOnly = $true; QuickTimeout = $true; ErrorAction = 'SilentlyContinue' }
        if ($Server) { $p.Server = $Server }
        $answers = @(Resolve-DnsName @p | Where-Object { $_.Section -eq 'Answer' -and $_.Type -eq $c[1] })
        if (-not $answers) {
            [pscustomobject]@{ Domain = $d; Name = $name; Type = $c[1]; Value = '(none)'; TTL = ''; Purpose = $c[2]; Flag = '' }
            continue
        }
        foreach ($a in $answers) {
            $value = switch ($c[1]) {
                'MX'    { "$($a.Preference) $($a.NameExchange)" }
                'TXT'   { ($a.Strings -join '') }
                'CNAME' { $a.NameHost }
                'SRV'   { "$($a.Priority) $($a.Weight) $($a.Port) $($a.NameTarget)" }
            }
            $flag = if ($a.TTL -gt $MaxTtl) { "Lower TTL (currently $($a.TTL)s)" } else { '' }
            [pscustomobject]@{ Domain = $d; Name = $name; Type = $c[1]; Value = $value; TTL = $a.TTL; Purpose = $c[2]; Flag = $flag }
        }
    }
}

$file = ".\DomainDns-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
$rows | Export-Csv $file -NoTypeInformation
$rows | Format-Table Domain, Type, Name, Value, TTL, Flag -AutoSize -Wrap | Out-Host
$high = @($rows | Where-Object Flag)
if ($high) { Write-Host "$($high.Count) record(s) have a TTL above $MaxTtl seconds. Lower them at least one old-TTL period before cutover." -ForegroundColor Yellow }
Write-Host "Saved to $file. Keep it: it's your rollback reference." -ForegroundColor Cyan
