<#
.SYNOPSIS
    Adds the moving domains to the destination tenant, waits for the verification record to reach the authoritative
    name servers, verifies, and prints the DNS records with what to do about each one, DKIM included.

.DESCRIPTION
    TL;DR: Run it BEFORE you release the domain from the source. It adds the domain unverified and gives you the
    MS=ms TXT record. Publish it, run with -WaitForDns to confirm every authoritative name server serves it, then
    run Remove-SourceDomain -RemoveDomain in the source, then this again with -Verify. On a live cutover verification
    was instant once the source let go, so the bounce window was minutes, not hours.

    Every record it prints gets an Action, compared against what the authoritative name server serves now:
      Unchanged   MX, SPF, autodiscover, enterpriseregistration, enterpriseenrollment already match Microsoft 365.
                  They are the same in every tenant (MX is <domain-with-dashes>.mail.protection.outlook.com).
      Change      The verification TXT and the DKIM CNAMEs. These are the records that are tenant specific.
      Add         Anything Microsoft wants that isn't there yet.
      Skip        Skype for Business legacy records (sip, lyncdiscover, _sip._tls, _sipfederationtls._tcp, msoid).
                  Teams doesn't need them.
      DoNotApply  Any apex (@) CNAME, for example to <tenant>.sharepoint.com. An apex CNAME breaks MX, TXT and
                  everything else at the root of the zone. Never apply it.
      Check       The lookup failed; compare by hand.
    After DKIM CNAMEs exist, run with -EnableDkim. Safe to re-run at every stage. Supports -WhatIf.

    Tip: the TXT can sit alongside the old tenant's MS= record and the SPF record. A zone can hold many TXT values.
    Trap: "verified" in the portal but Exchange doesn't know the domain yet: wait a few minutes before -EnableDkim.
    Paid tool: domain move features in migration suites add and verify for you; you still publish the TXT.
    No script: Microsoft 365 admin centre > Settings > Domains > Add domain. It shows the TXT and, later, the records.
    It also offers the Skype records and the apex SharePoint CNAME; ignore both.

.PARAMETER Domain
    Domains to add, e.g. fabrikam.com.

.PARAMETER TenantId
    Destination tenant ID (GUID). Required.

.PARAMETER AdminUpn
    Optional destination admin account. Open Exchange or Graph sessions for another account or tenant are
    disconnected first, and the script signs in again as this account.

.PARAMETER WaitForDns
    After adding, ask every authoritative name server for the MS= TXT every 30 seconds until all of them serve it,
    or -WaitMinutes passes. Goes straight on to -Verify if you gave that too.

.PARAMETER WaitMinutes
    How long -WaitForDns waits. Default 30.

.PARAMETER Verify
    Ask Microsoft 365 to check the verification TXT records. Only succeeds once the source has released the domain.

.PARAMETER EnableDkim
    Turn on DKIM signing. Only works once both selector CNAMEs resolve.

.EXAMPLE
    .\Add-TargetDomain.ps1 -TenantId <destination-guid> -AdminUpn admin@contoso.onmicrosoft.com -Domain fabrikam.com

.EXAMPLE
    .\Add-TargetDomain.ps1 -TenantId <destination-guid> -Domain fabrikam.com -WaitForDns

.EXAMPLE
    .\Add-TargetDomain.ps1 -TenantId <destination-guid> -Domain fabrikam.com -WaitForDns -Verify

.EXAMPLE
    .\Add-TargetDomain.ps1 -TenantId <destination-guid> -Domain fabrikam.com -EnableDkim

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.1.0
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: Microsoft.Graph.Authentication, Microsoft.Graph.Identity.DirectoryManagement, ExchangeOnlineManagement 3.x.
              Global Administrator or Domain Name Administrator plus Exchange Administrator in the destination.
              Windows (Resolve-DnsName) for -WaitForDns and the Action comparison; elsewhere Action shows Check.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][string[]]$Domain,
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F-]{36}$')][string]$TenantId,
    [string]$AdminUpn,
    [switch]$WaitForDns,
    [ValidateRange(1, 240)][int]$WaitMinutes = 30,
    [switch]$Verify,
    [switch]$EnableDkim
)

$ErrorActionPreference = 'Stop'
$Domain = @($Domain | ForEach-Object { $_.Trim().ToLower() })

# ------------------------------------------------------------ DNS helpers (authoritative, not a cache)
$canDns = [bool](Get-Command Resolve-DnsName -ErrorAction SilentlyContinue)
function Get-AuthNameServer([string]$Zone) {
    if (-not $canDns) { return @() }
    try { @(Resolve-DnsName -Name $Zone -Type NS -DnsOnly -ErrorAction Stop | Where-Object { $_.Section -eq 'Answer' -and $_.Type -eq 'NS' } | ForEach-Object { "$($_.NameHost)".TrimEnd('.') } | Sort-Object -Unique) }
    catch { @() }
}
function Get-DnsValue([string]$Name, [string]$Type, [string]$Server) {
    # Returns an array of values (empty if the record doesn't exist) or $null if the lookup itself failed
    if (-not $canDns) { return $null }
    $p = @{ Name = $Name; Type = $Type; DnsOnly = $true; QuickTimeout = $true; ErrorAction = 'Stop' }
    if ($Server) { $p.Server = $Server }
    try { $answers = @(Resolve-DnsName @p | Where-Object { $_.Section -eq 'Answer' -and "$($_.Type)" -eq $Type }) }
    catch {
        if ("$($_.Exception.Message)" -match 'does not exist|NXDOMAIN|No records|not found') { return , @() }
        return $null
    }
    $out = foreach ($r in $answers) {
        switch ($Type) {
            'MX'    { "$($r.Preference) $("$($r.NameExchange)".TrimEnd('.'))".ToLower() }
            'TXT'   { ($r.Strings -join '') }
            'CNAME' { "$($r.NameHost)".TrimEnd('.').ToLower() }
            'SRV'   { "$($r.Priority) $($r.Weight) $($r.Port) $("$($r.NameTarget)".TrimEnd('.'))".ToLower() }
            default { "$($r.IPAddress)" }
        }
    }
    , @($out)
}
function Get-RecordAction([string]$Type, [string]$Name, [string]$Value, $Current) {
    # $Current: array of current values at that name and type, or $null when the lookup failed
    $n = $Name.ToLower()
    if ($Type -eq 'CNAME' -and $n -eq '@') { return 'DoNotApply' }
    if ($n -in 'sip', 'lyncdiscover', 'msoid', '_sip._tls', '_sipfederationtls._tcp') { return 'Skip' }
    if ($null -eq $Current) { return 'Check' }
    $cur = @($Current | ForEach-Object { "$_".ToLower() })
    $v = "$Value".ToLower().TrimEnd('.')
    if ($Type -eq 'TXT' -and $v -like 'v=spf1*') {
        $spf = @($cur | Where-Object { $_ -like 'v=spf1*' })
        if ($spf | Where-Object { $_ -match 'include:spf\.protection\.outlook\.com' }) { return 'Unchanged' }
        if ($spf) { return 'Change' }
        return 'Add'
    }
    if ($Type -eq 'MX') {
        $host_ = ($v -split ' ')[-1]
        if ($cur | Where-Object { ($_ -split ' ')[-1] -eq $host_ }) { return 'Unchanged' }
        if ($cur) { return 'Change' }
        return 'Add'
    }
    if ($cur -contains $v) { return 'Unchanged' }
    if ($Type -eq 'TXT' -and $v -like 'ms=*') { return 'Change' }
    if ($n -like 'selector?._domainkey') { return 'Change' }
    if ($cur) { return 'Change' }
    'Add'
}

# ------------------------------------------------------------ connect, as the named account
function Connect-ExoTenant([string]$TenantId, [string]$AdminUpn) {
    $WhatIfPreference = $false
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

$ctx = Connect-GraphTenant $TenantId $AdminUpn @('Domain.ReadWrite.All')
$exo = Connect-ExoTenant $TenantId $AdminUpn
Write-Host "Signed in as $($ctx.Account) (Graph) and $($exo.UserPrincipalName) (Exchange) to tenant $TenantId" -ForegroundColor Cyan

$records = [System.Collections.Generic.List[object]]::new()
$script:ns = @()
function Add-Record($Domain, $Stage, $Type, $Name, $Value, $Note) {
    $fqdn = if ($Name -eq '@') { $Domain } else { "$Name.$Domain" }
    $current = Get-DnsValue $fqdn $Type ($script:ns | Select-Object -First 1)
    $action = Get-RecordAction $Type $Name $Value $current
    $records.Add([pscustomobject]@{ Domain = $Domain; Stage = $Stage; Action = $action; Type = $Type; Name = $Name; Value = $Value; Current = (@($current) -join ' | '); Note = $Note })
}

foreach ($d in $Domain) {
    Write-Host ''
    Write-Host "== $d" -ForegroundColor Cyan
    $script:ns = @(Get-AuthNameServer $d)
    if ($script:ns) { Write-Host "  DNS is hosted on: $($script:ns -join ', ')" -ForegroundColor Gray }
    $dom = Get-MgDomain -DomainId $d -ErrorAction SilentlyContinue

    # ---------------------------------------------------- add
    if (-not $dom) {
        if ($PSCmdlet.ShouldProcess($d, 'Add domain (unverified)')) {
            try { $dom = New-MgDomain -BodyParameter @{ id = $d }; Write-Host '  Added (unverified).' -ForegroundColor Green }
            catch { Write-Host "  Couldn't add: $($_.Exception.Message)" -ForegroundColor Red; continue }
        }
        else { Write-Host '  Would add it.' -ForegroundColor Cyan; continue }
    }

    # ---------------------------------------------------- verification record, wait, verify
    if (-not $dom.IsVerified) {
        $txt = Get-MgDomainVerificationDnsRecord -DomainId $d | Where-Object RecordType -eq 'Txt' | Select-Object -First 1
        $value = $txt.AdditionalProperties.text
        Add-Record $d 'Verify' 'TXT' '@' $value "TTL $($txt.Ttl). Sits alongside SPF and any old MS= record"
        Write-Host "  Verification: TXT @  $value" -ForegroundColor Yellow

        if ($WaitForDns) {
            if (-not $script:ns) { Write-Host "  Can't find the authoritative name servers (or no Resolve-DnsName here). Skipping the wait." -ForegroundColor Yellow }
            else {
                $deadline = (Get-Date).AddMinutes($WaitMinutes)
                do {
                    $missing = @($script:ns | Where-Object { $vals = Get-DnsValue $d 'TXT' $_; -not (@($vals) -contains $value) })
                    if (-not $missing) { Write-Host "  All $($script:ns.Count) name servers serve the TXT." -ForegroundColor Green; break }
                    Write-Host "  $(Get-Date -Format 'HH:mm:ss') Not yet on: $($missing -join ', '). Checking again in 30s." -ForegroundColor Gray
                    Start-Sleep -Seconds 30
                } while ((Get-Date) -lt $deadline)
                if ($missing) { Write-Host "  Gave up after $WaitMinutes min. Check the record was saved in the right zone, at the host the NS records point to." -ForegroundColor Red; continue }
            }
        }
        if ($Verify) {
            if ($PSCmdlet.ShouldProcess($d, 'Verify')) {
                try { $dom = Confirm-MgDomain -DomainId $d; Write-Host '  Verified.' -ForegroundColor Green }
                catch {
                    Write-Host "  Not verified yet: $($_.Exception.Message)" -ForegroundColor Red
                    Write-Host '  Check the TXT resolves on every name server, and that the source has finished removing the domain.' -ForegroundColor Yellow
                    continue
                }
            }
            else { continue }
        }
        else {
            Write-Host '  Publish that record. Then: -WaitForDns, Remove-SourceDomain -RemoveDomain in the source, then -Verify here.' -ForegroundColor Yellow
            continue
        }
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
        $type = "$($r.RecordType)"
        if ($type -eq 'Mx') { Add-Record $d 'Mail' 'MX' $label "$($p.preference) $($p.mailExchange)" 'Same in every tenant' }
        elseif ($type -eq 'Txt') { Add-Record $d 'Mail' 'TXT' $label $p.text 'SPF. Keep third-party senders if you change it' }
        elseif ($type -eq 'CName') {
            $note = if ($label -eq '@') { 'Never apply an apex CNAME: it breaks MX and TXT at the root' } else { "$($r.SupportedService)" }
            Add-Record $d 'Services' 'CNAME' $label $p.canonicalName $note
        }
        elseif ($type -eq 'Srv') { Add-Record $d 'Services' 'SRV' "$($p.service).$($p.protocol)" "$($p.priority) $($p.weight) $($p.port) $($p.nameTarget)" "$($r.SupportedService)" }
    }

    # ---------------------------------------------------- DKIM
    $accepted = Get-AcceptedDomain -Identity $d -ErrorAction SilentlyContinue
    if (-not $accepted) { Write-Host '  Exchange hasn''t picked the domain up yet. Re-run in a few minutes for DKIM.' -ForegroundColor Yellow; continue }
    $dkim = Get-DkimSigningConfig -Identity $d -ErrorAction SilentlyContinue
    if (-not $dkim -and $PSCmdlet.ShouldProcess($d, 'Create DKIM keys (disabled)')) {
        try { $dkim = New-DkimSigningConfig -DomainName $d -Enabled $false } catch { Write-Host "  DKIM: $($_.Exception.Message)" -ForegroundColor Red }
    }
    if ($dkim) {
        Add-Record $d 'DKIM' 'CNAME' 'selector1._domainkey' $dkim.Selector1CNAME 'Tenant specific: replace the old one'
        Add-Record $d 'DKIM' 'CNAME' 'selector2._domainkey' $dkim.Selector2CNAME 'Tenant specific: replace the old one'
        if ($EnableDkim -and -not $dkim.Enabled) {
            if ($PSCmdlet.ShouldProcess($d, 'Enable DKIM signing')) {
                try { Set-DkimSigningConfig -Identity $d -Enabled $true | Out-Null; Write-Host '  DKIM enabled.' -ForegroundColor Green }
                catch { Write-Host "  DKIM not enabled yet: $($_.Exception.Message). Check both CNAMEs resolve, then retry." -ForegroundColor Yellow }
            }
        }
        elseif ($dkim.Enabled) { Write-Host '  DKIM enabled.' -ForegroundColor Green }
    }
    $records.Add([pscustomobject]@{ Domain = $d; Stage = 'DMARC'; Action = 'Unchanged'; Type = 'TXT'; Name = '_dmarc'; Value = '(keep your existing record)'; Current = ''
            Note = 'If there is none, start with v=DMARC1; p=none; rua=mailto:<reports address>' })
}

Write-Host ''
$records | Sort-Object Domain, @{ e = { @{ Change = 0; Add = 1; Check = 2; DoNotApply = 3; Unchanged = 4; Skip = 5 }[$_.Action] } } |
    Format-Table Domain, Action, Type, Name, Value, Current -AutoSize -Wrap | Out-String -Width 250 | Write-Host
if ($records | Where-Object Action -eq 'DoNotApply') { Write-Host 'DoNotApply: never create an apex CNAME. It would take down mail and verification for the whole domain.' -ForegroundColor Red }
$file = ".\TargetDomainDns-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
$records | Export-Csv $file -NoTypeInformation -WhatIf:$false
Write-Host "Records saved to $file. Apply only Change and Add rows (Set-DomainDnsRecords.ps1 can do it through the provider's API)." -ForegroundColor Cyan
