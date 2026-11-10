<#
.SYNOPSIS
    Creates or updates the destination mail users that Microsoft's native cross-tenant mailbox migration moves
    mailboxes into, stamped with the source mailbox's ExchangeGuid, ArchiveGuid and X500 addresses.

.DESCRIPTION
    TL;DR: Native cross-tenant mailbox migration doesn't move a mailbox into a mailbox. It moves it into a mail user
    in the destination that carries the source mailbox's identity. If any of those attributes is missing or wrong,
    the move for that person fails. This script does the stamping in one pass:
      1. Source identity. Either reads a CSV that already has ExchangeGuid, ArchiveGuid, LegacyExchangeDN and
         EmailAddresses, or, with -SourceAdminUpn, signs in to the source and reads them live with Get-EXOMailbox,
         writing SourceMailboxIdentity-<date>.csv. 01's 04-Mailboxes.csv from version 1.1.0 holds all of these (older inventories don't),
         so the first run is normally with -SourceAdminUpn.
      2. Destination. Signs in to the destination and, for every source mailbox in the mapping file, creates a mail
         user (or mail-enables an existing user, or updates an existing mail user) with:
           -ExchangeGuid, -ArchiveGuid (only when the source archive is active),
           x500:<source LegacyExchangeDN> plus every x500 the source mailbox already has,
           primary SMTP and sign-in name = the mapping's Destination,
           ExternalEmailAddress = the source mailbox's <tenant>.onmicrosoft.com address.
      3. Writes a report CSV and a migration batch CSV (EmailAddress column) for New-MigrationBatch.
    It never assigns a licence. License the mail user after the move (or, at the earliest, after it's stamped):
    an Exchange licence on an unstamped object creates an empty mailbox and the move fails.

.PARAMETER MappingFile
    Source,Destination CSV from 02b (old address to new sign-in). Only rows whose Source is a source mailbox are used.
    Lists and contacts in the same file are ignored.

.PARAMETER SourceCsv
    Optional. A CSV with PrimarySmtpAddress, ExchangeGuid, ArchiveGuid, ArchiveStatus, LegacyExchangeDN and
    EmailAddresses (or X500Addresses), for example SourceMailboxIdentity-<date>.csv from an earlier run. 01's
    04-Mailboxes.csv on its own isn't enough; with -SourceAdminUpn, its PrimarySmtpAddress column limits which
    mailboxes are read.

.PARAMETER SourceAdminUpn
    Reads the source live. Signs in to the source as this account (closing any other Exchange Online session),
    reads every mapped mailbox, writes SourceMailboxIdentity-<date>.csv and disconnects.

.PARAMETER ExportOnly
    With -SourceAdminUpn: read the source and write the identity CSV, then stop. Nothing is changed anywhere.

.PARAMETER AdminUpn
    Destination admin. The script disconnects any Exchange Online session for another account, signs in as this one
    and prints "Signed in as X to tenant Y" before it changes anything.

.PARAMETER Only
    Only these source addresses. Use it for the first test batch.

.PARAMETER Force
    Overwrite an existing mail user whose ExchangeGuid is set to a different value. Without it, that row is marked
    Conflict and left alone.

.PARAMETER OutputFolder
    Where the identity, report and batch CSVs go. Default: the current folder.

.EXAMPLE
    .\New-TargetMailUsers.ps1 -MappingFile .\UserMapping.csv -SourceAdminUpn admin@fabrikam.onmicrosoft.com -ExportOnly
    Reads the source only and writes SourceMailboxIdentity-<date>.csv. Changes nothing.

.EXAMPLE
    .\New-TargetMailUsers.ps1 -MappingFile .\UserMapping.csv -SourceCsv .\SourceMailboxIdentity-20261110-1400.csv `
        -AdminUpn admin@contoso.onmicrosoft.com -Only alex@fabrikam.com, kim@fabrikam.com -WhatIf
    Preview the destination changes for a two-person test batch.

.EXAMPLE
    .\New-TargetMailUsers.ps1 -MappingFile .\UserMapping.csv -SourceAdminUpn admin@fabrikam.onmicrosoft.com `
        -AdminUpn admin@contoso.onmicrosoft.com
    Read the source live, then create and stamp every mapped mail user in the destination.

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.0.1
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: ExchangeOnlineManagement 3.x. Source: View-Only Recipients or higher. Destination: Mail Recipients or
              Organization Management. Built from Microsoft's documented target-object prerequisites; not yet run
              end to end in my lab.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][string]$MappingFile,
    [string]$SourceCsv,
    [string]$SourceAdminUpn,
    [switch]$ExportOnly,
    [string]$AdminUpn,
    [string[]]$Only,
    [switch]$Force,
    [string]$OutputFolder = '.'
)

$ErrorActionPreference = 'Stop'
$stamp = Get-Date -Format 'yyyyMMdd-HHmm'
$EmptyGuid = '00000000-0000-0000-0000-000000000000'
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -WhatIf:$false | Out-Null }

# ------------------------------------------------------------ helpers
function Test-RealGuid([string]$g) { $g -and $g -ne $EmptyGuid -and $g -match '^[0-9a-fA-F-]{36}$' }

function Connect-ExoAs {
    # Opens Exchange Online as $Upn, closing any session for another account first. Returns the connection.
    param([string]$Upn, [string]$Label)
    $open = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object State -eq 'Connected' | Select-Object -First 1
    if ($open -and $Upn -and "$($open.UserPrincipalName)" -ne $Upn) {
        Write-Host "Exchange Online is signed in as $($open.UserPrincipalName). Reconnecting as $Upn." -ForegroundColor Yellow
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        $open = $null
    }
    if (-not $open) {
        $a = @{ ShowBanner = $false }
        if ($Upn) { $a.UserPrincipalName = $Upn }
        Connect-ExchangeOnline @a
        $open = Get-ConnectionInformation | Where-Object State -eq 'Connected' | Select-Object -First 1
    }
    if ($Upn -and "$($open.UserPrincipalName)" -ne $Upn) { throw "$Label`: Exchange Online signed in as $($open.UserPrincipalName), not $Upn." }
    Write-Host "Signed in as $($open.UserPrincipalName) to tenant $($open.TenantID) ($Label)" -ForegroundColor Cyan
    $open
}

function Get-X500List($addresses) {
    @($addresses | ForEach-Object { "$_" } | Where-Object { $_ -match '^x500:' } | ForEach-Object { 'x500:' + $_.Substring(5) })
}

function Get-RoutingAddress($addresses) {
    # The source mailbox's <tenant>.onmicrosoft.com address (not .mail.onmicrosoft.com), for ExternalEmailAddress.
    $all = @($addresses | ForEach-Object { "$_" } | Where-Object { $_ -match '^smtp:' } | ForEach-Object { $_.Substring(5) })
    $hit = $all | Where-Object { $_ -match '\.onmicrosoft\.com$' -and $_ -notmatch '\.mail\.onmicrosoft\.com$' } | Select-Object -First 1
    if (-not $hit) { $hit = $all | Where-Object { $_ -match '\.mail\.onmicrosoft\.com$' } | Select-Object -First 1 }
    $hit
}

# ------------------------------------------------------------ mapping
$mapRows = @(Import-Csv $MappingFile)
foreach ($c in 'Source', 'Destination') { if (-not ($mapRows | Get-Member -Name $c)) { throw "$MappingFile needs a '$c' column." } }
$map = @{}
foreach ($r in $mapRows) { if ($r.Source -and $r.Destination) { $map[$r.Source.Trim().ToLower()] = $r.Destination.Trim() } }
if ($Only) {
    $keep = @($Only | ForEach-Object { $_.Trim().ToLower() })
    foreach ($k in @($map.Keys)) { if ($keep -notcontains $k) { $map.Remove($k) } }
    foreach ($o in $keep) { if (-not $map.ContainsKey($o)) { Write-Warning "$o isn't in $MappingFile (or has no Destination)." } }
}
if ($map.Count -eq 0) { throw 'No mapping rows to work on.' }

# ------------------------------------------------------------ 1. source identity
$identity = @()
$csvRows = if ($SourceCsv) { @(Import-Csv $SourceCsv) } else { @() }
$csvHasIdentity = $csvRows.Count -gt 0 -and ($csvRows | Get-Member -Name ExchangeGuid) -and ($csvRows | Get-Member -Name LegacyExchangeDN)

if ($SourceAdminUpn) {
    $null = Connect-ExoAs -Upn $SourceAdminUpn -Label 'source'
    $wanted = @($map.Keys)
    if ($csvRows.Count -and ($csvRows | Get-Member -Name PrimarySmtpAddress)) {
        $inCsv = @($csvRows | ForEach-Object { "$($_.PrimarySmtpAddress)".ToLower() })
        $wanted = @($wanted | Where-Object { $inCsv -contains $_ })
    }
    $identity = foreach ($addr in $wanted) {
        try {
            $m = Get-EXOMailbox -Identity $addr -Properties ExchangeGuid, ArchiveGuid, ArchiveStatus, LegacyExchangeDN, EmailAddresses, RecipientTypeDetails -ErrorAction Stop
        }
        catch { Write-Verbose "$addr is not a source mailbox: $($_.Exception.Message)"; continue }
        [pscustomobject]@{
            PrimarySmtpAddress   = "$($m.PrimarySmtpAddress)"
            DisplayName          = $m.DisplayName
            Alias                = $m.Alias
            RecipientTypeDetails = "$($m.RecipientTypeDetails)"
            ExchangeGuid         = "$($m.ExchangeGuid)"
            ArchiveGuid          = "$($m.ArchiveGuid)"
            ArchiveStatus        = "$($m.ArchiveStatus)"
            LegacyExchangeDN     = $m.LegacyExchangeDN
            RoutingAddress       = Get-RoutingAddress $m.EmailAddresses
            X500Addresses        = (Get-X500List $m.EmailAddresses) -join ';'
        }
    }
    $identity = @($identity)
    $idFile = Join-Path $OutputFolder "SourceMailboxIdentity-$stamp.csv"
    $identity | Export-Csv $idFile -NoTypeInformation -WhatIf:$false
    Write-Host ("Source: {0} mailboxes read, written to {1}" -f $identity.Count, $idFile) -ForegroundColor Green
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    if ($ExportOnly) { return }
}
elseif ($csvHasIdentity) {
    $identity = foreach ($r in $csvRows) {
        $x500 = if ($r.PSObject.Properties['X500Addresses']) { "$($r.X500Addresses)" -split ';' | Where-Object { $_ } }
                elseif ($r.PSObject.Properties['EmailAddresses']) { Get-X500List ("$($r.EmailAddresses)" -split '[;\s]+') }
        $routing = if ($r.PSObject.Properties['RoutingAddress'] -and $r.RoutingAddress) { $r.RoutingAddress }
                   elseif ($r.PSObject.Properties['EmailAddresses']) { Get-RoutingAddress ("$($r.EmailAddresses)" -split '[;\s]+') }
        [pscustomobject]@{
            PrimarySmtpAddress   = $r.PrimarySmtpAddress
            DisplayName          = $r.DisplayName
            Alias                = $(if ($r.PSObject.Properties['Alias']) { $r.Alias })
            RecipientTypeDetails = $r.RecipientTypeDetails
            ExchangeGuid         = $r.ExchangeGuid
            ArchiveGuid          = $r.ArchiveGuid
            ArchiveStatus        = $r.ArchiveStatus
            LegacyExchangeDN     = $r.LegacyExchangeDN
            RoutingAddress       = $routing
            X500Addresses        = (@($x500) -join ';')
        }
    }
    $identity = @($identity)
}
elseif ($ExportOnly) { throw '-ExportOnly needs -SourceAdminUpn.' }
else {
    throw ("{0} has no ExchangeGuid and LegacyExchangeDN columns (01's 04-Mailboxes.csv has them from version 1.1.0; re-run the inventory if yours is older). " -f $(if ($SourceCsv) { $SourceCsv } else { 'No -SourceCsv was given, and it' })) +
          'Run again with -SourceAdminUpn <source admin> to read them live.'
}

# ------------------------------------------------------------ 2. destination
$conn = Connect-ExoAs -Upn $AdminUpn -Label 'destination'
$accepted = @(Get-AcceptedDomain | ForEach-Object { "$($_.DomainName)".ToLower() })

# Wrong-tenant guard: every Destination domain must be accepted here, and the source GUIDs mustn't be local mailboxes.
$destDomains = @($map.Values | ForEach-Object { $_.Split('@')[-1].ToLower() } | Sort-Object -Unique)
$missing = @($destDomains | Where-Object { $accepted -notcontains $_ })
if ($missing.Count -eq $destDomains.Count) {
    throw "None of the Destination domains ($($destDomains -join ', ')) is an accepted domain in tenant $($conn.TenantID). Wrong tenant? Use -AdminUpn."
}

$report = [System.Collections.Generic.List[object]]::new()
$batch = [System.Collections.Generic.List[object]]::new()
$byAddr = @{}
foreach ($i in $identity) { $byAddr["$($i.PrimarySmtpAddress)".ToLower()] = $i }

foreach ($src in ($map.Keys | Sort-Object)) {
    $dest = $map[$src]
    $row = [ordered]@{ Source = $src; Destination = $dest; Type = ''; Action = ''; ExchangeGuid = ''; ArchiveGuid = ''; X500Count = 0; ExternalEmailAddress = ''; Note = '' }
    $i = $byAddr[$src]
    if (-not $i) { continue }   # not a source mailbox (a list, contact or leaver): nothing to do here
    $row.Type = $i.RecipientTypeDetails
    $row.ExchangeGuid = $i.ExchangeGuid
    $hasArchive = $i.ArchiveStatus -eq 'Active' -and (Test-RealGuid $i.ArchiveGuid)
    if ($hasArchive) { $row.ArchiveGuid = $i.ArchiveGuid }

    $x500 = [System.Collections.Generic.List[string]]::new()
    if ($i.LegacyExchangeDN) { $x500.Add("x500:$($i.LegacyExchangeDN)") }
    foreach ($x in ("$($i.X500Addresses)" -split ';' | Where-Object { $_ })) { if (-not ($x500 -contains $x)) { $x500.Add($x) } }
    $row.X500Count = $x500.Count
    $row.ExternalEmailAddress = $i.RoutingAddress

    $problems = @()
    if (-not (Test-RealGuid $i.ExchangeGuid)) { $problems += 'source ExchangeGuid missing' }
    if (-not $i.LegacyExchangeDN) { $problems += 'source LegacyExchangeDN missing' }
    if (-not $i.RoutingAddress) { $problems += 'source has no onmicrosoft.com address for ExternalEmailAddress' }
    if ($accepted -notcontains $dest.Split('@')[-1].ToLower()) { $problems += "$($dest.Split('@')[-1]) isn't an accepted domain in the destination" }
    if ($i.RecipientTypeDetails -match 'Room|Equipment|Scheduling') { $problems += "$($i.RecipientTypeDetails) mailboxes: check Microsoft's current support before including" }
    if ($problems) { $row.Action = 'Skip'; $row.Note = $problems -join '; '; $report.Add([pscustomobject]$row); continue }

    try {
        $existing = Get-Recipient -Identity $dest -ErrorAction SilentlyContinue
        if (-not $existing) { $existing = Get-User -Identity $dest -ErrorAction SilentlyContinue }
        $kind = if ($existing) { "$($existing.RecipientTypeDetails)" } else { '' }

        if ($kind -match 'Mailbox$') {
            $row.Action = 'Conflict'
            $row.Note = "Already a $kind in the destination. The move would fail. Remove the mailbox (and its licence) or use another copy option for this person"
            $report.Add([pscustomobject]$row); continue
        }

        if ($kind -eq 'MailUser') {
            $mu = Get-MailUser -Identity $dest
            if ((Test-RealGuid "$($mu.ExchangeGuid)") -and "$($mu.ExchangeGuid)" -ne $i.ExchangeGuid -and -not $Force) {
                $row.Action = 'Conflict'; $row.Note = "Mail user already has ExchangeGuid $($mu.ExchangeGuid). Use -Force to overwrite"
                $report.Add([pscustomobject]$row); continue
            }
            $row.Action = 'Update'
        }
        elseif ($kind) {
            # A user with no mail attributes: mail-enable it rather than create a duplicate.
            $row.Action = 'Enable'
            if ($PSCmdlet.ShouldProcess($dest, "Enable-MailUser -ExternalEmailAddress $($i.RoutingAddress)")) {
                Enable-MailUser -Identity $dest -ExternalEmailAddress $i.RoutingAddress -PrimarySmtpAddress $dest | Out-Null
            }
        }
        else {
            $row.Action = 'Create'
            $name = if ($i.DisplayName) { $i.DisplayName } else { $dest.Split('@')[0] }
            $alias = $dest.Split('@')[0] -replace '[^A-Za-z0-9!#$%&''*+\-/=?^_`{|}~.]', ''
            # Random password nobody knows. 07 sets the temporary password people actually use at cutover.
            $pw = ConvertTo-SecureString ([Convert]::ToBase64String((1..24 | ForEach-Object { [byte](Get-Random -Maximum 256) })) + 'aA1!') -AsPlainText -Force
            if ($PSCmdlet.ShouldProcess($dest, "New-MailUser -ExternalEmailAddress $($i.RoutingAddress)")) {
                New-MailUser -Name $name -DisplayName $name -Alias $alias -MicrosoftOnlineServicesID $dest -PrimarySmtpAddress $dest `
                    -ExternalEmailAddress $i.RoutingAddress -Password $pw | Out-Null
            }
        }

        # Stamp. PrimarySmtpAddress is set again with ExternalEmailAddress, or Exchange makes the external address primary.
        $set = @{ Identity = $dest; ExchangeGuid = $i.ExchangeGuid; ExternalEmailAddress = $i.RoutingAddress; PrimarySmtpAddress = $dest
                  EmailAddresses = @{ Add = @($x500) } }
        if ($hasArchive) { $set.ArchiveGuid = $i.ArchiveGuid }
        if ($PSCmdlet.ShouldProcess($dest, "Set-MailUser ExchangeGuid $($i.ExchangeGuid)$(if ($hasArchive) { ", ArchiveGuid $($i.ArchiveGuid)" }), $($x500.Count) x500")) {
            if ($kind -eq '' -or $row.Action -eq 'Enable') { Start-Sleep -Seconds 2 }   # let a new object settle
            Set-MailUser @set -ErrorAction Stop
            $batch.Add([pscustomobject]@{ EmailAddress = $dest })
        }
        else { $row.Action = "WhatIf:$($row.Action)" }
    }
    catch {
        $row.Action = 'Error'; $row.Note = $_.Exception.Message
    }
    $report.Add([pscustomobject]$row)
}

$notMailbox = $map.Count - $report.Count
$repFile = Join-Path $OutputFolder "TargetMailUsers-Report-$stamp.csv"
$report | Export-Csv $repFile -NoTypeInformation -WhatIf:$false
if ($batch.Count) {
    $batchFile = Join-Path $OutputFolder "MigrationBatch-$stamp.csv"
    $batch | Export-Csv $batchFile -NoTypeInformation -WhatIf:$false
    Write-Host "Migration batch CSV: $batchFile (EmailAddress = destination address, as New-MigrationBatch wants)" -ForegroundColor Green
}
Write-Host ''
$report | Group-Object Action | Select-Object Name, Count | Format-Table -AutoSize | Out-Host
Write-Host "Report: $repFile. $notMailbox mapping rows weren't source mailboxes and were left out." -ForegroundColor Green
if ($report | Where-Object Action -in 'Skip', 'Conflict', 'Error') { Write-Host 'Fix every Skip, Conflict and Error before you start a batch for that person.' -ForegroundColor Yellow }
Write-Host "Don't license these mail users until they're stamped. Safest: license after the move completes." -ForegroundColor Yellow
