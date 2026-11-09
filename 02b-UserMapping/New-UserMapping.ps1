<#
.SYNOPSIS
    Builds the user mapping file (old address to new sign-in) by reading both tenants and matching every recipient,
    and writes it as a CSV for this toolkit plus a .sgum file for ShareGate.

.DESCRIPTION
    TL;DR: Every later step (permissions, forwarding, lists, cutover, validation, and the copy tool's user mapping)
    depends on one file that says who is who. Built by hand it's slow and one typo puts a shared mailbox's permissions
    on the wrong person. This script signs in to the source, then the destination, reads every mailbox, list and
    contact, and matches them in this order:

      1. Overrides file          your own decisions, always win (blank Destination = leave this one out)
      2. Build plan              part 02's TargetBuildPlan.csv: SourceAddress -> TargetAddress
      3. Toolkit tag             objects 05 and 06 created carry the source address in CustomAttribute15
      4. Old address             the destination already has the source address as an alias
      5. Alias                   same alias (with -AliasPrefix too), one match only
      6. Address local part      alex@fabrikam.com -> alex@contoso.com, one match only
      7. Display name            one match only, and flagged for review

    Guest accounts are never a match, and a mailbox only ever maps to a mailbox. If the destination tenant also hosts
    other companies, pass -DestinationDomain so steps 5 to 7 can't pick someone else's 'info@' or 'closings@'.

    Nothing in either tenant is changed. You get:
      UserMapping.csv          Source,Destination (+ type, method). Matched rows only. This is the file 05, 06, 07, 09 use
      UserMapping.sgum         The same people for ShareGate's user and group mapping (mailbox users only, by sign-in)
      UserMapping-Review.csv   Every source recipient, how it matched, and what needs a human look

.PARAMETER SourceAdmin
    Optional admin sign-in for the source tenant, to pre-fill the sign-in box.

.PARAMETER DestinationAdmin
    Optional admin sign-in for the destination tenant.

.PARAMETER OverridesFile
    Optional CSV with Source,Destination. Wins over every automatic match. A blank Destination leaves the source out
    of the mapping (leavers you aren't moving, test accounts).

.PARAMETER PlanFile
    Optional. Part 02's TargetBuildPlan.csv. Its SourceAddress -> TargetAddress rows (Include = True) are used before
    any guessing, which matters when staged mailboxes have a different alias (a prefix, a renamed shared mailbox).

.PARAMETER DestinationDomain
    Optional. Destination domains this migration owns (staging domain, onmicrosoft domain). Steps 5 to 7 only match
    destination objects on these domains. Use it whenever the destination tenant has other companies' mailboxes.

.PARAMETER AliasPrefix
    Optional. A prefix you gave staged objects in the destination, e.g. 'fab-'. Alias and local part matching also try
    '<prefix><alias>', and prefer it.

.PARAMETER Exclude
    Source addresses or wildcards to leave out, e.g. 'healthmailbox*','discoverysearch*'.

.PARAMETER CompareWith
    An existing mapping CSV (Source,Destination). Reports every row where the new mapping disagrees with it. Useful if
    you already built one by hand.

.PARAMETER OutputFolder
    Where to write the files. Default: current folder.

.EXAMPLE
    .\New-UserMapping.ps1 -SourceAdmin admin@fabrikam.com -DestinationAdmin admin@contoso.com

.EXAMPLE
    .\New-UserMapping.ps1 -OverridesFile .\MappingOverrides.csv -CompareWith .\UserMapping-old.csv

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.1.0
    Date    : 2026-11-09
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: ExchangeOnlineManagement 3.x, read access (View-Only Recipients or higher) in both tenants.
    Sign-in names change at cutover. Re-run this after 07 so the .sgum points at the final sign-ins before any
    copy-tool incremental.
#>
[CmdletBinding()]
param(
    [string]$SourceAdmin,
    [string]$DestinationAdmin,
    [string]$OverridesFile,
    [string]$PlanFile,
    [string[]]$DestinationDomain,
    [string]$AliasPrefix,
    [string[]]$Exclude = @('healthmailbox*', 'discoverysearchmailbox*', 'systemmailbox*'),
    [string]$CompareWith,
    [string]$OutputFolder = '.'
)

$ErrorActionPreference = 'Stop'
if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) { throw 'Install-Module ExchangeOnlineManagement first.' }

function Get-SmtpOnly($a) { ("$a" -replace '^smtp:', '').ToLower() }

function Read-Tenant($Label, $Upn) {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    Write-Host "Sign in to the $Label tenant." -ForegroundColor Magenta
    $p = @{ ShowBanner = $false }
    if ($Upn) { $p.UserPrincipalName = $Upn }
    Connect-ExchangeOnline @p
    $tenantId = (Get-ConnectionInformation | Select-Object -Last 1).TenantID
    $name = (Get-AcceptedDomain | Where-Object Default).DomainName
    Write-Host "  Connected to $name. Reading recipients..." -ForegroundColor Cyan

    $upns = @{}
    Get-EXOMailbox -ResultSize Unlimited | ForEach-Object { $upns[$_.ExternalDirectoryObjectId] = $_.UserPrincipalName.ToLower() }
    $rows = Get-Recipient -ResultSize Unlimited |
        Where-Object { $_.PrimarySmtpAddress -and $_.RecipientTypeDetails -notmatch 'Discovery|Monitoring|Arbitration|AuditLog' } |
        ForEach-Object {
            $addr = "$($_.PrimarySmtpAddress)".ToLower()
            [pscustomobject]@{
                Address = $addr
                Local   = $addr.Split('@')[0]
                Alias   = "$($_.Alias)".ToLower()
                Name    = "$($_.DisplayName)".Trim().ToLower()
                Display = "$($_.DisplayName)"
                Type    = "$($_.RecipientTypeDetails)"
                Upn     = $upns["$($_.ExternalDirectoryObjectId)"]
                Tag     = "$($_.CustomAttribute15)"
                Smtp    = @($_.EmailAddresses | Where-Object { "$_" -like 'smtp:*' } | ForEach-Object { Get-SmtpOnly $_ })
            }
        }
    Write-Host "  $(@($rows).Count) recipients." -ForegroundColor Cyan
    [pscustomobject]@{ TenantId = $tenantId; Name = $name; Rows = @($rows) }
}

$src = Read-Tenant 'SOURCE' $SourceAdmin
$dst = Read-Tenant 'DESTINATION' $DestinationAdmin
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
if ($src.TenantId -eq $dst.TenantId) { throw "Both sign-ins went to the same tenant ($($src.Name)). Sign out of the browser or use a private window and run again." }

# Indexes on the destination
$byTag = @{}; $byProxy = @{}; $byAlias = @{}; $byLocal = @{}; $byName = @{}
function Add-Index($h, $k, $v) { if ($k) { if (-not $h.ContainsKey($k)) { $h[$k] = [System.Collections.Generic.List[object]]::new() }; $h[$k].Add($v) } }
foreach ($d in $dst.Rows) {
    if ($d.Tag -like 'T2T:*') { Add-Index $byTag $d.Tag.Substring(4).ToLower() $d }
    foreach ($s in $d.Smtp) { Add-Index $byProxy $s $d }
    Add-Index $byAlias $d.Alias $d
    Add-Index $byLocal $d.Local $d
    Add-Index $byName $d.Name $d
}
$dstByKey = @{}
foreach ($d in $dst.Rows) { $dstByKey[$d.Address] = $d; if ($d.Upn) { $dstByKey[$d.Upn] = $d } }

$overrides = @{}
if ($OverridesFile) {
    foreach ($o in Import-Csv $OverridesFile) { $overrides["$($o.Source)".Trim().ToLower()] = "$($o.Destination)".Trim().ToLower() }
    Write-Host "$($overrides.Count) override(s) loaded." -ForegroundColor Cyan
}

$domains = @($DestinationDomain | Where-Object { $_ } | ForEach-Object { $_.ToLower().TrimStart('@') })
function Test-Fit($s, $d, [switch]$Scoped) {
    if ($d.Type -match 'Guest') { return $false }                                   # a guest is never the new home
    if (($s.Type -match 'Mailbox$' -and $s.Type -ne 'GroupMailbox') -and -not ($d.Type -match 'Mailbox$' -and $d.Type -ne 'GroupMailbox')) { return $false }
    if ($Scoped -and $domains) { return [bool](@($d.Address, $d.Upn) | Where-Object { $_ -and $domains -contains $_.Split('@')[-1] }) }
    $true
}
function Get-One($h, $k, $s, [switch]$Scoped) {
    if ($k -and $h.ContainsKey($k)) {
        $l = @($h[$k] | Where-Object { Test-Fit $s $_ -Scoped:$Scoped } | Sort-Object Address -Unique)
        if ($l.Count -eq 1) { return $l[0] } elseif ($l.Count -gt 1) { return 'many' }
    }
}

$plan = @{}
if ($PlanFile) {
    foreach ($p in Import-Csv $PlanFile | Where-Object { $_.Include -eq 'True' }) { $plan["$($p.SourceAddress)".Trim().ToLower()] = "$($p.TargetAddress)".Trim().ToLower() }
    Write-Host "$($plan.Count) row(s) from the build plan." -ForegroundColor Cyan
}
$destAccepted = @($dst.Rows | ForEach-Object { $_.Address.Split('@')[-1] } | Sort-Object -Unique)
$unscoped = (-not $domains -and $destAccepted.Count -gt 3)
if ($unscoped) {
    Write-Warning "The destination has addresses on $($destAccepted.Count) domains. If it hosts other companies, re-run with -DestinationDomain so alias and name matches can't land on their mailboxes."
}

$review = foreach ($s in $src.Rows) {
    $match = $null; $method = ''; $note = ''
    if (@($Exclude | Where-Object { $s.Address -like $_ })) { $method = 'Excluded' }
    elseif ($overrides.ContainsKey($s.Address)) {
        $t = $overrides[$s.Address]
        if (-not $t) { $method = 'Excluded (override)' }
        elseif ($dstByKey.ContainsKey($t)) { $match = $dstByKey[$t]; $method = 'Override' }
        else { $method = 'Override'; $note = "Override target $t not found in the destination" }
    }
    elseif ($plan.ContainsKey($s.Address)) {
        $t = $plan[$s.Address]
        $hit = if ($dstByKey.ContainsKey($t)) { $dstByKey[$t] } else { Get-One $byProxy $t $s }
        if ($hit -and $hit -ne 'many') { $match = $hit; $method = 'Build plan' }
        else { $method = 'Build plan'; $note = "Plan target $t not found in the destination" }
    }
    else {
        $p = "$AliasPrefix".ToLower()
        foreach ($step in @(
                @{ M = 'Toolkit tag'; F = { Get-One $byTag $s.Address $s } },
                @{ M = 'Old address'; F = { foreach ($a in @($s.Address) + $s.Smtp) { $r = Get-One $byProxy $a $s; if ($r) { return $r } } } },
                @{ M = 'Alias'; F = { if ($p) { $r = Get-One $byAlias "$p$($s.Alias)" $s -Scoped; if ($r) { return $r } }; Get-One $byAlias $s.Alias $s -Scoped } },
                @{ M = 'Local part'; F = { if ($p) { $r = Get-One $byLocal "$p$($s.Local)" $s -Scoped; if ($r) { return $r } }; Get-One $byLocal $s.Local $s -Scoped } },
                @{ M = 'Display name'; F = { Get-One $byName $s.Name $s -Scoped } })) {
            $r = & $step.F
            if ($r -eq 'many') { $note = "$($step.M) matches more than one destination object"; continue }
            if ($r) { $match = $r; $method = $step.M; break }
        }
        if (-not $match -and -not $note) { $note = 'No match' }
    }
    $needsLook = (-not $match -and $method -notlike 'Excluded*') -or $method -eq 'Display name' -or [bool]$note
    if ($match -and $method -eq 'Display name') { $note = 'Matched on name only. Check it''s the same person' }
    elseif ($match -and $unscoped -and $method -in 'Alias', 'Local part') { $needsLook = $true; $note = "Guessed by $($method.ToLower()) in a tenant with several domains. Confirm, or re-run with -DestinationDomain" }
    if ($match -and ($s.Type -match 'Mailbox') -ne ($match.Type -match 'Mailbox')) { $needsLook = $true; $note = (@($note, "Type differs: $($s.Type) -> $($match.Type)") | Where-Object { $_ }) -join '. ' }
    [pscustomobject]@{
        Source          = $s.Address
        SourceName      = $s.Display
        SourceType      = $s.Type
        SourceUpn       = $s.Upn
        Destination     = if ($match) { if ($match.Upn) { $match.Upn } else { $match.Address } } else { '' }
        DestinationName = if ($match) { $match.Display } else { '' }
        DestinationType = if ($match) { $match.Type } else { '' }
        DestinationUpn  = if ($match) { $match.Upn } else { '' }
        Method          = if ($method) { $method } else { 'None' }
        Review          = $needsLook
        Note            = $note
    }
}

# Two sources landing on one destination is almost always a mistake
$review | Where-Object Destination | Group-Object Destination | Where-Object Count -gt 1 | ForEach-Object {
    foreach ($r in $_.Group) { $r.Review = $true; $r.Note = (@($r.Note, "Shares destination with $($_.Count - 1) other source(s)") | Where-Object { $_ }) -join '. ' }
}

New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
$csv = Join-Path $OutputFolder 'UserMapping.csv'
$rev = Join-Path $OutputFolder 'UserMapping-Review.csv'
$sg = Join-Path $OutputFolder 'UserMapping.sgum'
$mapped = @($review | Where-Object Destination)
$mapped | Select-Object Source, Destination, SourceType, DestinationType, Method | Export-Csv $csv -NoTypeInformation
$review | Sort-Object @{e = { -not $_.Review } }, Method, Source | Export-Csv $rev -NoTypeInformation

# ShareGate wants sign-in claims, so only rows with a sign-in on both sides
$x = [System.Text.StringBuilder]::new()
[void]$x.AppendLine('<?xml version="1.0" encoding="utf-8"?>')
[void]$x.AppendLine('<UserAndGroupMappings xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">')
[void]$x.AppendLine('  <Mappings>')
$sgCount = 0
foreach ($r in $mapped | Where-Object { $_.SourceUpn -and $_.DestinationUpn }) {
    $s = [Security.SecurityElement]::Escape($r.SourceUpn); $d = [Security.SecurityElement]::Escape($r.DestinationUpn)
    [void]$x.AppendLine("    <Mapping><Source AccountName=""i:0#.f|membership|$s"" DisplayName=""$s"" PrincipalType=""User"" /><Destination AccountName=""i:0#.f|membership|$d"" DisplayName=""$d"" PrincipalType=""User"" /></Mapping>")
    $sgCount++
}
[void]$x.AppendLine('  </Mappings>')
[void]$x.AppendLine('</UserAndGroupMappings>')
[IO.File]::WriteAllText((Join-Path (Resolve-Path $OutputFolder).Path 'UserMapping.sgum'), $x.ToString(), [Text.UTF8Encoding]::new($false))

Write-Host ''
$review | Group-Object Method | Sort-Object Count -Descending | Format-Table @{n = 'Matched by'; e = { $_.Name } }, Count -AutoSize | Out-String -Width 250 | Write-Host
$look = @($review | Where-Object Review)
Write-Host "$($mapped.Count) mapped, $sgCount in the ShareGate file, $($look.Count) to review." -ForegroundColor $(if ($look) { 'Yellow' } else { 'Green' })
if ($look) { $look | Format-Table Source, SourceType, Destination, Method, Note -AutoSize -Wrap | Out-String -Width 250 | Write-Host }

if ($CompareWith) {
    $old = @{}; foreach ($o in Import-Csv $CompareWith) { $old["$($o.Source)".Trim().ToLower()] = "$($o.Destination)".Trim().ToLower() }
    $diff = foreach ($k in $old.Keys) {
        $new = ($review | Where-Object Source -eq $k | Select-Object -First 1).Destination
        if ("$new" -ne $old[$k]) { [pscustomobject]@{ Source = $k; Old = $old[$k]; New = $new } }
    }
    $extra = @($mapped | Where-Object { -not $old.ContainsKey($_.Source) })
    Write-Host "Compared with ${CompareWith}: $(@($diff).Count) disagree, $($extra.Count) new in this run." -ForegroundColor Cyan
    if ($diff) { $diff | Format-Table -AutoSize | Out-String -Width 250 | Write-Host }
}
Write-Host "Written: $csv, $sg, $rev. These hold names and addresses: keep them out of source control." -ForegroundColor Cyan
