<#
.SYNOPSIS
    Cutover in the destination: gives every migrated recipient its old addresses back, adds X500 addresses so replies
    to old mail don't bounce, switches sign-in names to the real domain, and optionally sets temporary passwords.

.DESCRIPTION
    TL;DR: Run after Add-TargetDomain has verified the domains. For each row in your mapping file (Source,Destination)
    and each list or contact Copy-MailConfig created, it adds the recipient's old addresses on the moved domains and
    an X500 address built from its old LegacyExchangeDN. For recipients in TargetBuildPlan.csv it also makes the old
    address primary and the sign-in name. Everything else, for example people who already lived in the destination,
    just gets the old addresses as aliases. Safe to re-run. Supports -WhatIf.

    Why X500: migrated mail, meeting invites and Outlook's autocomplete address people by their old internal Exchange
    address, not their email address. Without a matching X500 address, replying to an old message bounces with an
    IMCEAEX non-delivery report.

    Before it changes anything it checks the snapshot is the -pre one (it holds addresses on the domain and
    LegacyExchangeDN values). A snapshot taken after Remove-SourceDomain stripped the addresses has neither, and on a
    live cutover that silently cost everyone their aliases and X500s. It stops unless you add -Force.

    It also checks every old address against the destination first. If a guest, mail user, contact or someone
    else's mailbox already holds it, that address is logged as Conflict and skipped. Guests are the usual culprit:
    when people from the source were invited into the destination as guests, the guest holds their old address.
    -RemoveConflictingGuests deletes those guests (they go to the recycle bin for 30 days) and retries.

    Matching is strictly by the mapping file's Destination column. Same-name mailboxes on other domains, for
    example another company hosted in the same destination tenant, are never touched.

    Tip: run it with -WhatIf first and read every Conflict and Problem line.
    Paid tool: most migration tools can add X500s during the final pass if you feed them the LegacyExchangeDN. Check
    the tool does it for groups and contacts too; many only do mailboxes.
    No script: Exchange admin centre > Recipients > the mailbox > Email addresses > add the old address and an X500
    "address type" entry by hand. Fine for five people.

.PARAMETER TenantId
    Destination tenant ID (GUID). Required.

.PARAMETER MappingFile
    CSV with Source and Destination columns: old address to destination recipient (UPN or address). Add rows for
    Microsoft 365 groups and Teams too, e.g. sales@fabrikam.com -> sales@contoso.onmicrosoft.com.

.PARAMETER SourceSnapshot
    SourceRecipients-pre-<stamp>.csv from Remove-SourceDomain. Holds the old addresses and LegacyExchangeDN values.
    Use the night's file: Remove-SourceDomain -SnapshotOnly run on cutover night just before the release (the
    -RemoveDomain run's own -pre file works too). The T-1 day file is the fallback. Never a -post file.

.PARAMETER Domain
    The domains that moved. Only addresses on these (and accepted in the destination) are added.

.PARAMETER PlanFile
    TargetBuildPlan.csv from New-TargetMailboxes. Recipients in it (Include True) get the full switch: primary address
    and sign-in name. Default .\TargetBuildPlan.csv.

.PARAMETER ResetPasswords
    Give every person (Kind User in the plan) a new temporary password, and write them to a CSV with the new sign-in
    names. They're asked to change it at first sign-in.

.PARAMETER SkipUpn
    Change addresses but leave sign-in names alone.

.PARAMETER AdminUpn
    Optional destination admin account. Open Exchange or Graph sessions for another account or tenant are
    disconnected first, and the script signs in again as this account.

.PARAMETER Force
    Carry on even if the snapshot looks like it was taken after the source addresses were removed.

.PARAMETER RemoveConflictingGuests
    When a destination guest (GuestMailUser) already holds an old address, delete that guest through Graph (it goes
    to the recycle bin) and retry the add. Never touches mailboxes, mail users or contacts. Supports -WhatIf.

.EXAMPLE
    .\Switch-TargetAddresses.ps1 -TenantId <destination-guid> -MappingFile .\UserMapping.csv -SourceSnapshot .\SourceRecipients-pre-<stamp>.csv `
        -Domain fabrikam.com -WhatIf

.EXAMPLE
    .\Switch-TargetAddresses.ps1 -TenantId <destination-guid> -AdminUpn admin@contoso.onmicrosoft.com -MappingFile .\UserMapping.csv `
        -SourceSnapshot .\SourceRecipients-pre-<stamp>.csv -Domain fabrikam.com, fabrikam.co.uk -RemoveConflictingGuests -ResetPasswords

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.1.1
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: ExchangeOnlineManagement 3.x, Microsoft.Graph.Authentication, Microsoft.Graph.Users. Global Administrator
              (or User Administrator plus Exchange Administrator) in the destination.
    Changing a sign-in name renames the person's OneDrive URL and signs them out of Office apps and Teams.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F-]{36}$')][string]$TenantId,
    [Parameter(Mandatory)][string]$MappingFile,
    [Parameter(Mandatory)][string]$SourceSnapshot,
    [Parameter(Mandatory)][string[]]$Domain,
    [string]$PlanFile = '.\TargetBuildPlan.csv',
    [switch]$ResetPasswords,
    [switch]$SkipUpn,
    [string]$AdminUpn,
    [switch]$Force,
    [switch]$RemoveConflictingGuests
)

$ErrorActionPreference = 'Stop'
$Domain = @($Domain | ForEach-Object { $_.Trim().ToLower() })

function Get-DomainOf($a) { "$a".Split('@')[-1].ToLower() }

# ------------------------------------------------------------ snapshot sanity, before connecting to anything
if (-not (Test-Path $SourceSnapshot)) { throw "Snapshot $SourceSnapshot not found." }
if (-not (Test-Path $MappingFile)) { throw "Mapping file $MappingFile not found." }
$snapRows = @(Import-Csv $SourceSnapshot)
$mapRows = @(Import-Csv $MappingFile | Where-Object { $_.Source -and $_.Destination })
$withDomain = @($snapRows | Where-Object { "$($_.EmailAddresses)".Split(';') | Where-Object { $_ -match '^smtp:(.+)$' -and $Domain -contains (Get-DomainOf $Matches[1]) } }).Count
$withLegDn = @($snapRows | Where-Object { $_.LegacyExchangeDN }).Count
$snapAddr = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($row in $snapRows) {
    foreach ($a in "$($row.EmailAddresses)".Split(';')) { if ($a -match '^smtp:(.+)$') { [void]$snapAddr.Add($Matches[1]) } }
    [void]$snapAddr.Add("$($row.PrimarySmtpAddress)")
}
$mapped = @($mapRows | Where-Object { $Domain -contains (Get-DomainOf $_.Source.Trim()) })
$found = @($mapped | Where-Object { $snapAddr.Contains($_.Source.Trim()) }).Count
$pct = if ($mapped.Count) { [math]::Round(100 * $found / $mapped.Count) } else { 100 }
$stageCol = @($snapRows | Select-Object -First 1 -ExpandProperty SnapshotStage -ErrorAction SilentlyContinue)
Write-Host "Snapshot: $($snapRows.Count) rows, $withDomain with addresses on $($Domain -join ', '), $withLegDn with LegacyExchangeDN; $found of $($mapped.Count) mapped sources found ($pct%)." -ForegroundColor Cyan
if ($withDomain -eq 0 -or $withLegDn -eq 0 -or $pct -lt 20 -or $stageCol -contains 'AfterRelease') {
    $msg = 'This snapshot was taken after the source addresses were removed. Use the -pre snapshot.'
    if (-not $Force) { throw $msg }
    Write-Warning "$msg Carrying on because of -Force: expect 'Not in the source snapshot' lines and no X500s."
}

# ------------------------------------------------------------ connect and guard, as the named account
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
$exo = Connect-ExoTenant $TenantId $AdminUpn
$ctx = Connect-GraphTenant $TenantId $AdminUpn @('User.ReadWrite.All', 'User-PasswordProfile.ReadWrite.All')
Write-Host "Signed in as $($exo.UserPrincipalName) (Exchange) and $($ctx.Account) (Graph) to tenant $TenantId" -ForegroundColor Cyan

$accepted = @((Get-AcceptedDomain).DomainName | ForEach-Object { $_.ToLower() })
$usable = @($Domain | Where-Object { $accepted -contains $_ })
foreach ($d in $Domain | Where-Object { $usable -notcontains $_ }) {
    Write-Warning "$d isn't an accepted domain here yet. Addresses on it are skipped. Run Add-TargetDomain -Verify first."
}
if (-not $usable) { throw 'None of the domains are accepted in this tenant yet.' }
Write-Host "Destination confirmed. Switching addresses on: $($usable -join ', ')" -ForegroundColor Cyan

# ------------------------------------------------------------ inputs
$snapIndex = @{}
foreach ($s in $snapRows) {
    foreach ($a in "$($s.EmailAddresses)".Split(';')) {
        if ($a -match '^smtp:(.+)$') { $snapIndex[$Matches[1].ToLower()] = $s }
    }
    $snapIndex["$($s.PrimarySmtpAddress)".ToLower()] = $s
}
$plan = @{}
if (Test-Path $PlanFile) {
    Import-Csv $PlanFile | Where-Object { $_.Include -eq 'True' } | ForEach-Object { $plan["$($_.SourceAddress)".ToLower()] = $_ }
}
else { Write-Warning "$PlanFile not found. Every recipient gets aliases only: no primary address or sign-in changes." }

$work = [System.Collections.Generic.List[object]]::new()
$seen = [System.Collections.Generic.HashSet[string]]::new()
foreach ($m in $mapRows) {
    if ($m.Source -and $m.Destination -and $seen.Add($m.Source.Trim().ToLower())) {
        $work.Add([pscustomobject]@{ Source = $m.Source.Trim().ToLower(); Destination = $m.Destination.Trim(); Origin = 'Mapping' })
    }
}
# Lists and contacts Copy-MailConfig created carry their source address in CustomAttribute15
foreach ($r in Get-Recipient -Filter "CustomAttribute15 -like 'T2T:*'" -ResultSize Unlimited -ErrorAction SilentlyContinue) {
    $src = "$($r.CustomAttribute15)".Substring(4).ToLower()
    if ($seen.Add($src)) { $work.Add([pscustomobject]@{ Source = $src; Destination = "$($r.PrimarySmtpAddress)"; Origin = 'Tagged' }) }
}
Write-Host "$($work.Count) recipients to process." -ForegroundColor Cyan

$log = [System.Collections.Generic.List[object]]::new()
$passwords = [System.Collections.Generic.List[object]]::new()
function Write-Log($Object, $Status, $Detail) {
    $log.Add([pscustomobject]@{ Source = $Object; Status = $Status; Detail = $Detail })
    $c = @{ Done = 'Green'; Exists = 'Gray'; WhatIf = 'Cyan'; Skipped = 'Yellow'; Problem = 'Red'; Conflict = 'Magenta' }[$Status]
    Write-Host ("  [{0}] {1}: {2}" -f $Status, $Object, $Detail) -ForegroundColor $c
}
function New-RandomPassword {
    $rng = [System.Security.Cryptography.RandomNumberGenerator]
    $sets = 'ABCDEFGHJKLMNPQRSTUVWXYZ', 'abcdefghijkmnpqrstuvwxyz', '23456789', '!@#$%*?'
    $all = -join $sets
    $chars = [System.Collections.Generic.List[char]]::new()
    foreach ($s in $sets) { $chars.Add($s[$rng::GetInt32($s.Length)]) }
    1..12 | ForEach-Object { $chars.Add($all[$rng::GetInt32($all.Length)]) }
    -join ($chars | Sort-Object { $rng::GetInt32([int]::MaxValue) })
}

function Get-ConflictingRecipient($Address, $Dest) {
    # Returns the recipient holding $Address if it isn't $Dest, else $null
    $r = @(Get-EXORecipient -Identity $Address -Properties ExternalDirectoryObjectId, RecipientTypeDetails -ErrorAction SilentlyContinue)
    if (-not $r) { $r = @(Get-Recipient -Identity $Address -ErrorAction SilentlyContinue) }
    foreach ($x in $r) {
        $same = ($x.ExternalDirectoryObjectId -and "$($x.ExternalDirectoryObjectId)" -eq "$($Dest.ExternalDirectoryObjectId)") -or ("$($x.Guid)" -and "$($x.Guid)" -eq "$($Dest.Guid)")
        if (-not $same) { return $x }
    }
    $null
}
function Wait-AddressFree($Address, $Dest) {
    # Directory deletes take a minute or two to reach Exchange
    for ($i = 0; $i -lt 6; $i++) {
        Start-Sleep -Seconds 30
        if (-not (Get-ConflictingRecipient $Address $Dest)) { return $true }
    }
    $false
}

# ------------------------------------------------------------ per recipient
foreach ($w in $work) {
    # Strictly the mapping file's Destination: never a lookup by display name
    $dest = Get-EXORecipient -Identity $w.Destination -Properties EmailAddresses, ExternalDirectoryObjectId, RecipientTypeDetails -ErrorAction SilentlyContinue
    if (-not $dest) { Write-Log $w.Source 'Problem' "Destination $($w.Destination) not found"; continue }
    $type = "$($dest.RecipientTypeDetails)"
    $set, $primaryParam = switch -Regex ($type) {
        '^GroupMailbox$'                          { 'Set-UnifiedGroup', 'PrimarySmtpAddress'; break }
        'Mailbox$'                                { 'Set-Mailbox', 'WindowsEmailAddress'; break }
        '^Mail(Universal|NonUniversal)'           { 'Set-DistributionGroup', 'PrimarySmtpAddress'; break }
        '^DynamicDistributionGroup$'              { 'Set-DynamicDistributionGroup', 'PrimarySmtpAddress'; break }
        '^(MailUser|GuestMailUser)$'              { 'Set-MailUser', 'PrimarySmtpAddress'; break }
        '^MailContact$'                           { 'Set-MailContact', $null; break }
        default                                   { $null, $null }
    }
    if (-not $set) { Write-Log $w.Source 'Problem' "Recipient type $type not handled"; continue }

    $have = @($dest.EmailAddresses | ForEach-Object { "$_".ToLower() })
    $snap = $snapIndex[$w.Source]
    $add = [System.Collections.Generic.List[string]]::new()
    if ($snap) {
        if ($type -ne 'MailContact') {
            foreach ($a in "$($snap.EmailAddresses)".Split(';')) {
                if ($a -match '^smtp:(.+)$' -and $usable -contains (Get-DomainOf $Matches[1])) { $add.Add("smtp:$($Matches[1].ToLower())") }
            }
        }
        if ($snap.LegacyExchangeDN) { $add.Add("X500:$($snap.LegacyExchangeDN)") }
        foreach ($a in "$($snap.EmailAddresses)".Split(';')) { if ($a -match '^x500:') { $add.Add("X500:$($a.Substring(5))") } }
    }
    elseif ($usable -contains (Get-DomainOf $w.Source) -and $type -ne 'MailContact') {
        $add.Add("smtp:$($w.Source)")
        Write-Log $w.Source 'Skipped' 'Not in the source snapshot, so no X500 address. Adding the address only'
    }
    $add = @($add | Select-Object -Unique | Where-Object { $have -notcontains $_.ToLower() })

    # Conflicts: an old address already on a different object in the destination
    $sourceBlocked = $false
    foreach ($a in @($add | Where-Object { $_ -like 'smtp:*' })) {
        $addr = $a.Substring(5)
        $other = Get-ConflictingRecipient $addr $dest
        if (-not $other) { continue }
        $what = "$($other.RecipientTypeDetails) $($other.DisplayName) ($($other.ExternalDirectoryObjectId))"
        if ($RemoveConflictingGuests -and "$($other.RecipientTypeDetails)" -eq 'GuestMailUser') {
            if ($PSCmdlet.ShouldProcess($what, "Delete guest holding $addr (recycle bin)")) {
                try {
                    Remove-MgUser -UserId $other.ExternalDirectoryObjectId
                    Write-Log $w.Source 'Done' "Deleted guest $what that held $addr"
                    if (Wait-AddressFree $addr $dest) { continue }
                    Write-Log $w.Source 'Conflict' "$addr still shows on the deleted guest after 3 minutes. Run again shortly"
                }
                catch { Write-Log $w.Source 'Problem' "Couldn't delete guest $($what): $($_.Exception.Message)" }
            }
            else { Write-Log $w.Source 'WhatIf' "Would delete guest $what that holds $addr, then add it"; continue }
        }
        else {
            $hint = if ("$($other.RecipientTypeDetails)" -eq 'GuestMailUser') { ' Rerun with -RemoveConflictingGuests, or delete the guest in Entra.' } else { ' Remove it from that object by hand if it should move.' }
            Write-Log $w.Source 'Conflict' "$addr is already used by $what. Skipped.$hint"
        }
        $add = @($add | Where-Object { $_ -ne $a })
        if ($addr -eq $w.Source) { $sourceBlocked = $true }
    }

    $full = $plan.ContainsKey($w.Source) -or ($w.Origin -eq 'Tagged' -and $type -ne 'MailContact')
    $makePrimary = $full -and $primaryParam -and -not $sourceBlocked -and ($usable -contains (Get-DomainOf $w.Source)) -and
                   ("$($dest.PrimarySmtpAddress)".ToLower() -ne $w.Source)
    $id = if ($dest.ExternalDirectoryObjectId) { $dest.ExternalDirectoryObjectId } else { "$($dest.Identity)" }

    try {
        if ($add) {
            if ($PSCmdlet.ShouldProcess($w.Destination, "Add $($add.Count) address(es)")) {
                & $set -Identity $id -EmailAddresses @{ Add = $add } -ErrorAction Stop
                Write-Log $w.Source 'Done' "Added $($add -join ', ')"
            }
            else { Write-Log $w.Source 'WhatIf' "Would add $($add -join ', ')" }
        }
        if ($makePrimary) {
            if ($PSCmdlet.ShouldProcess($w.Destination, "Primary address -> $($w.Source)")) {
                $p = @{ Identity = $id; ErrorAction = 'Stop' }; $p[$primaryParam] = $w.Source
                & $set @p
                Write-Log $w.Source 'Done' "Primary address now $($w.Source)"
            }
            else { Write-Log $w.Source 'WhatIf' "Would make $($w.Source) the primary address" }
        }
        if (-not $add -and -not $makePrimary -and -not $sourceBlocked) { Write-Log $w.Source 'Exists' 'Addresses already in place' }
    }
    catch { Write-Log $w.Source 'Problem' $_.Exception.Message; continue }

    # ---------------------------------------------------- sign-in name and password
    $isAccount = $type -match '^(User|Shared|Room|Equipment)Mailbox$'
    if ($full -and $isAccount -and -not $SkipUpn -and -not $sourceBlocked -and ($usable -contains (Get-DomainOf $w.Source))) {
        try {
            $u = Get-MgUser -UserId $dest.ExternalDirectoryObjectId -Property Id, UserPrincipalName
            if ($u.UserPrincipalName.ToLower() -ne $w.Source) {
                $clash = Get-MgUser -Filter "userPrincipalName eq '$($w.Source)'" -ErrorAction SilentlyContinue
                if ($clash) { Write-Log $w.Source 'Problem' "Sign-in $($w.Source) already belongs to another account" }
                elseif ($PSCmdlet.ShouldProcess($u.UserPrincipalName, "Sign-in -> $($w.Source)")) {
                    Update-MgUser -UserId $u.Id -UserPrincipalName $w.Source
                    Write-Log $w.Source 'Done' "Sign-in changed from $($u.UserPrincipalName)"
                }
                else { Write-Log $w.Source 'WhatIf' "Would change sign-in from $($u.UserPrincipalName)" }
            }
        }
        catch { Write-Log $w.Source 'Problem' "Sign-in: $($_.Exception.Message)" }
    }
    if ($ResetPasswords -and $plan.ContainsKey($w.Source) -and $plan[$w.Source].Kind -eq 'User') {
        if ($PSCmdlet.ShouldProcess($w.Source, 'Set a temporary password')) {
            try {
                $pw = New-RandomPassword
                Update-MgUser -UserId $dest.ExternalDirectoryObjectId -AccountEnabled:$true -PasswordProfile @{ Password = $pw; ForceChangePasswordNextSignIn = $true }
                $signIn = if ($SkipUpn) { (Get-MgUser -UserId $dest.ExternalDirectoryObjectId -Property UserPrincipalName).UserPrincipalName } else { $w.Source }
                $passwords.Add([pscustomobject]@{ DisplayName = $dest.DisplayName; SignIn = $signIn; TempPassword = $pw })
                Write-Log $w.Source 'Done' 'Temporary password set'
            }
            catch { Write-Log $w.Source 'Problem' "Password: $($_.Exception.Message)" }
        }
        else { Write-Log $w.Source 'WhatIf' 'Would set a temporary password' }
    }
}

# ------------------------------------------------------------ report
$stamp = Get-Date -Format 'yyyyMMdd-HHmm'
$log | Export-Csv ".\SwitchTargetAddresses-$stamp.csv" -NoTypeInformation -WhatIf:$false
if ($passwords.Count) {
    $passwords | Export-Csv ".\CutoverPasswords-$stamp.csv" -NoTypeInformation -WhatIf:$false
    Write-Warning "Temporary passwords saved to CutoverPasswords-$stamp.csv. Send them one by one, then delete the file."
}
Write-Host ''
$log | Group-Object Status | Sort-Object Name | ForEach-Object { Write-Host ("{0,-8} {1}" -f $_.Name, $_.Count) }
if ($log | Where-Object Status -eq 'Conflict') { Write-Host 'Conflicts: those addresses were not added. See the report for who holds each one.' -ForegroundColor Magenta }
Write-Host "Report: .\SwitchTargetAddresses-$stamp.csv" -ForegroundColor Cyan
Write-Host 'Sign-in changes reach OneDrive URLs, Teams and Office over the next hour or so. People will be asked to sign in again.' -ForegroundColor Yellow
