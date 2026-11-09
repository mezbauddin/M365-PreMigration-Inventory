<#
.SYNOPSIS
    Adds the moving domains to the destination tenant, verifies them, and prints the exact DNS records to set, DKIM
    included.

.DESCRIPTION
    TL;DR: Run it once to add the domains and get the verification TXT records. Put those in DNS, then run it again
    with -Verify. Once verified it prints the MX, SPF, autodiscover and DKIM records the destination wants (read from
    the tenant, not guessed) and saves them to a CSV. After the DKIM CNAMEs exist, run with -EnableDkim.
    Safe to re-run at every stage. Supports -WhatIf.

    A domain can't be verified here until the source tenant has finished removing it (Remove-SourceDomain).

.PARAMETER Domain
    Domains to add, e.g. fabrikam.com.

.PARAMETER TenantId
    Destination tenant ID (GUID). Required.

.PARAMETER Verify
    Ask Microsoft 365 to check the verification TXT records.

.PARAMETER EnableDkim
    Turn on DKIM signing. Only works once both selector CNAMEs resolve.

.EXAMPLE
    .\Add-TargetDomain.ps1 -TenantId <dest-guid> -Domain fabrikam.com

.EXAMPLE
    .\Add-TargetDomain.ps1 -TenantId <dest-guid> -Domain fabrikam.com -Verify

.EXAMPLE
    .\Add-TargetDomain.ps1 -TenantId <dest-guid> -Domain fabrikam.com -EnableDkim

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.0.0
    Date    : 2026-11-09
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: Microsoft.Graph.Authentication, Microsoft.Graph.Identity.DirectoryManagement, ExchangeOnlineManagement 3.x.
              Global Administrator or Domain Name Administrator plus Exchange Administrator in the destination.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][string[]]$Domain,
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F-]{36}$')][string]$TenantId,
    [switch]$Verify,
    [switch]$EnableDkim
)

$ErrorActionPreference = 'Stop'
$Domain = @($Domain | ForEach-Object { $_.Trim().ToLower() })

Connect-MgGraph -TenantId $TenantId -NoWelcome -Scopes 'Domain.ReadWrite.All'
if ((Get-MgContext).TenantId -ne $TenantId) { throw "Graph signed in to $((Get-MgContext).TenantId), not $TenantId." }
if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object State -eq 'Connected')) { Connect-ExchangeOnline -ShowBanner:$false }
$exoTenant = (Get-ConnectionInformation | Where-Object State -eq 'Connected' | Select-Object -First 1).TenantID
if ($exoTenant -ne $TenantId) { throw "Exchange Online is signed in to $exoTenant, not $TenantId. Disconnect-ExchangeOnline and sign in to the destination." }
Write-Host 'Destination tenant confirmed.' -ForegroundColor Cyan

$records = [System.Collections.Generic.List[object]]::new()
function Add-Record($Domain, $Stage, $Type, $Name, $Value, $Note) {
    $records.Add([pscustomobject]@{ Domain = $Domain; Stage = $Stage; Type = $Type; Name = $Name; Value = $Value; Note = $Note })
}

foreach ($d in $Domain) {
    Write-Host ''
    Write-Host "== $d" -ForegroundColor Cyan
    $dom = Get-MgDomain -DomainId $d -ErrorAction SilentlyContinue

    # ---------------------------------------------------- add
    if (-not $dom) {
        if ($PSCmdlet.ShouldProcess($d, 'Add domain')) {
            try { $dom = New-MgDomain -BodyParameter @{ id = $d }; Write-Host '  Added.' -ForegroundColor Green }
            catch { Write-Host "  Couldn't add: $($_.Exception.Message)" -ForegroundColor Red; continue }
        }
        else { Write-Host '  Would add it.' -ForegroundColor Cyan; continue }
    }

    # ---------------------------------------------------- verify
    if (-not $dom.IsVerified) {
        $txt = Get-MgDomainVerificationDnsRecord -DomainId $d | Where-Object RecordType -eq 'Txt' | Select-Object -First 1
        $value = $txt.AdditionalProperties.text
        Add-Record $d 'Verify' 'TXT' '@' $value "TTL $($txt.Ttl). Can sit alongside the existing SPF record"
        Write-Host "  Verification: TXT @  $value" -ForegroundColor Yellow
        if ($Verify) {
            if ($PSCmdlet.ShouldProcess($d, 'Verify')) {
                try { $dom = Confirm-MgDomain -DomainId $d; Write-Host '  Verified.' -ForegroundColor Green }
                catch {
                    Write-Host "  Not verified yet: $($_.Exception.Message)" -ForegroundColor Red
                    Write-Host '  Check the TXT record resolves publicly, and that the source has finished removing the domain.' -ForegroundColor Yellow
                    continue
                }
            }
            else { continue }
        }
        else { Write-Host '  Add that record, then run again with -Verify.' -ForegroundColor Yellow; continue }
    }
    else { Write-Host '  Verified.' -ForegroundColor Green }

    # ---------------------------------------------------- mail services
    $dom = Get-MgDomain -DomainId $d
    if ($dom.SupportedServices -notcontains 'Email') {
        if ($PSCmdlet.ShouldProcess($d, 'Enable Email service')) {
            Update-MgDomain -DomainId $d -SupportedServices (@($dom.SupportedServices) + 'Email' | Select-Object -Unique)
        }
    }
    foreach ($r in Get-MgDomainServiceConfigurationRecord -DomainId $d -All) {
        $p = $r.AdditionalProperties
        $label = if ($r.Label -eq $d) { '@' } else { $r.Label -replace "\.$([regex]::Escape($d))$", '' }
        switch ($r.RecordType) {
            'Mx'    { Add-Record $d 'Mail' 'MX' $label "$($p.preference) $($p.mailExchange)" 'Replace the old MX' }
            'Txt'   { Add-Record $d 'Mail' 'TXT' $label $p.text 'SPF. Replace the old v=spf1 record, keep any third-party senders in it' }
            'CName' { Add-Record $d 'Services' 'CNAME' $label $p.canonicalName $r.SupportedService }
            'Srv'   { Add-Record $d 'Services' 'SRV' "$($p.service).$($p.protocol)" "$($p.priority) $($p.weight) $($p.port) $($p.nameTarget)" $r.SupportedService }
        }
    }

    # ---------------------------------------------------- DKIM
    $accepted = Get-AcceptedDomain -Identity $d -ErrorAction SilentlyContinue
    if (-not $accepted) { Write-Host '  Exchange hasn''t picked the domain up yet. Re-run in a few minutes for DKIM.' -ForegroundColor Yellow; continue }
    $dkim = Get-DkimSigningConfig -Identity $d -ErrorAction SilentlyContinue
    if (-not $dkim -and $PSCmdlet.ShouldProcess($d, 'Create DKIM keys (disabled)')) {
        try { $dkim = New-DkimSigningConfig -DomainName $d -Enabled $false } catch { Write-Host "  DKIM: $($_.Exception.Message)" -ForegroundColor Red }
    }
    if ($dkim) {
        Add-Record $d 'DKIM' 'CNAME' 'selector1._domainkey' $dkim.Selector1CNAME 'Replace the old one'
        Add-Record $d 'DKIM' 'CNAME' 'selector2._domainkey' $dkim.Selector2CNAME 'Replace the old one'
        if ($EnableDkim -and -not $dkim.Enabled) {
            if ($PSCmdlet.ShouldProcess($d, 'Enable DKIM signing')) {
                try { Set-DkimSigningConfig -Identity $d -Enabled $true | Out-Null; Write-Host '  DKIM enabled.' -ForegroundColor Green }
                catch { Write-Host "  DKIM not enabled yet: $($_.Exception.Message). Check both CNAMEs resolve, then retry." -ForegroundColor Yellow }
            }
        }
        elseif ($dkim.Enabled) { Write-Host '  DKIM enabled.' -ForegroundColor Green }
    }
    Add-Record $d 'DMARC' 'TXT' '_dmarc' '(keep your existing record)' 'If there is none, start with v=DMARC1; p=none; rua=mailto:<reports address>'
}

Write-Host ''
$records | Format-Table Domain, Stage, Type, Name, Value -AutoSize -Wrap | Out-Host
$file = ".\TargetDomainDns-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
$records | Export-Csv $file -NoTypeInformation
Write-Host "Records saved to $file. Compare with your Export-DomainDns file and change only what's listed." -ForegroundColor Cyan
