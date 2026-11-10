<#
.SYNOPSIS
    Releases domains from the source tenant: snapshots every recipient, lists everything that would block removal,
    moves sign-in names and addresses off the domains, and removes the domains.

.DESCRIPTION
    TL;DR: A domain can only live in one tenant. Before the destination can verify it, the source has to let go.
    This script:
      1. Snapshots every recipient (addresses and LegacyExchangeDN) to SourceRecipients-pre-<date>.csv. The
         destination needs that file to add X500 addresses, so replies to old mail keep working.
      2. Lists every blocker, in -WhatIf and -BlockersOnly runs too: deleted users in the recycle bin, synced users,
         app registrations with identifier URIs on the domain (Application), Entra users with proxy addresses on the
         domain but no Exchange recipient (ProxyOnly), and anything else the directory says still references it.
         If a hard blocker is found on a real run, it stops before changing anything.
      3. With -FixApplications, moves app identifier URIs on the domain to the onmicrosoft.com domain.
      4. Moves the default domain to onmicrosoft.com if one of the domains is the default.
      5. Changes every sign-in name on the domains to <name>@<tenant>.onmicrosoft.com.
      6. Changes every recipient's primary address off the domains and removes their addresses on them. Where
         Exchange still holds an address as the WindowsLiveId (the sign-in change hasn't synced yet, a few
         minutes), it logs SyncLag and retries every minute for -RetryMinutes.
      7. With -RemoveDomain, disables DKIM and removes the domains.

    Best order on the night (it was the difference between minutes and hours of bounced mail on a live cutover):
    add the domain to the destination first with Add-TargetDomain (unverified), publish the MS=ms TXT record,
    confirm every authoritative name server serves it, then run this with -RemoveDomain, then Add-TargetDomain
    -Verify straight away. Mail to the domain bounces from the moment its addresses are stripped here until the
    destination verifies it.

    Tip: run -BlockersOnly a few days before. Every blocker it shows is one you don't want to meet at 9pm.
    Trap: a snapshot taken after the addresses are stripped is useless for X500s. Keep the -pre file safe.
    Paid tool: some migration suites have a domain move feature that does this and the destination side together.
    ShareGate doesn't touch domains.

.PARAMETER Domain
    Domains to release, e.g. fabrikam.com. Never the source's onmicrosoft.com domain.

.PARAMETER TenantId
    Source tenant ID (GUID). Required: the script refuses to touch any other tenant.

.PARAMETER AdminUpn
    Optional source admin account, e.g. admin@fabrikam.onmicrosoft.com. If an Exchange or Graph session is open as
    someone else, or to another tenant, it is disconnected and the script signs in again as this account.

.PARAMETER SnapshotOnly
    Write the recipient snapshot and stop. Run this first, today, so you have it whatever happens.

.PARAMETER BlockersOnly
    Read-only. List the blockers and stop. No snapshot, no changes. Use it in the days before cutover.

.PARAMETER FixApplications
    Change app registration identifier URIs on the domains to the tenant's onmicrosoft.com domain, keeping the path.
    Supports -WhatIf. Apps that other systems call by that URI need those systems updating too.

.PARAMETER RetryMinutes
    How long to keep retrying address removals blocked by WindowsLiveId sync lag, once a minute. Default 10.
    0 turns retrying off.

.PARAMETER RemoveDomain
    After moving everything off, remove the domains from the tenant.

.EXAMPLE
    .\Remove-SourceDomain.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -Domain fabrikam.com -BlockersOnly

.EXAMPLE
    .\Remove-SourceDomain.ps1 -TenantId <source-guid> -Domain fabrikam.com -SnapshotOnly

.EXAMPLE
    .\Remove-SourceDomain.ps1 -TenantId <source-guid> -Domain fabrikam.com, fabrikam.co.uk -WhatIf

.EXAMPLE
    .\Remove-SourceDomain.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -Domain fabrikam.com -FixApplications -RemoveDomain

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.1.0
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: ExchangeOnlineManagement 3.x, Microsoft.Graph.Authentication, Microsoft.Graph.Users,
              Microsoft.Graph.Applications, Microsoft.Graph.Identity.DirectoryManagement. Global Administrator in
              the source.
    Directory sync must be off for the source: synced users' sign-in names and addresses can't be changed in the cloud.
    No script: the Microsoft 365 admin centre, Settings > Domains > select the domain > Remove domain, runs a wizard
    that moves users, groups and proxy addresses for you. It is slower for big tenants, doesn't snapshot anything,
    and doesn't touch app registrations, but it does fix ProxyOnly users that Graph can't edit.
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)][string[]]$Domain,
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F-]{36}$')][string]$TenantId,
    [string]$AdminUpn,
    [switch]$SnapshotOnly,
    [switch]$BlockersOnly,
    [switch]$FixApplications,
    [ValidateRange(0, 120)][int]$RetryMinutes = 10,
    [switch]$RemoveDomain
)

$ErrorActionPreference = 'Stop'
$Domain = @($Domain | ForEach-Object { $_.Trim().ToLower() })
if ($Domain | Where-Object { $_ -like '*.onmicrosoft.com' }) { throw "Don't release an onmicrosoft.com domain." }

# ------------------------------------------------------------ connect, both to the same tenant, as the named account
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
$ctx = Connect-GraphTenant $TenantId $AdminUpn @('User.ReadWrite.All', 'Domain.ReadWrite.All', 'Directory.ReadWrite.All', 'Application.ReadWrite.All')
Write-Host "Signed in as $($exo.UserPrincipalName) (Exchange) and $($ctx.Account) (Graph) to tenant $TenantId" -ForegroundColor Cyan

$domains = @(Get-MgDomain -All)
$initial = ($domains | Where-Object IsInitial).Id
foreach ($d in $Domain) { if ($domains.Id -notcontains $d) { throw "$d isn't a domain in this tenant." } }
Write-Host "Source tenant confirmed. Fallback domain: $initial" -ForegroundColor Cyan

function Test-OnDomain($address) { $Domain -contains ("$address".Split('@')[-1].ToLower()) }
function Test-HostOnDomain($uri) {
    $h = try { ([uri]"$uri").Host.ToLower() } catch { '' }
    if (-not $h) { return $false }
    foreach ($d in $Domain) { if ($h -eq $d -or $h.EndsWith(".$d")) { return $true } }
    $false
}
$log = [System.Collections.Generic.List[object]]::new()
function Write-Log($Area, $Object, $Status, $Detail, $Fix = '') {
    $log.Add([pscustomobject]@{ Area = $Area; Object = $Object; Status = $Status; Detail = $Detail; Fix = $Fix })
    $c = @{ Done = 'Green'; WhatIf = 'Cyan'; Problem = 'Red'; Info = 'Gray'; Blocker = 'Yellow'; Warn = 'Yellow'; SyncLag = 'Magenta' }[$Status]
    $line = "  [{0}] {1}: {2} {3}" -f $Status, $Area, $Object, $Detail
    if ($Fix) { $line += "`n           Fix: $Fix" }
    Write-Host $line -ForegroundColor $c
}

# Get-EXORecipient cannot return LegacyExchangeDN; the classic cmdlet can
$recipients = @(Get-Recipient -ResultSize Unlimited)
$recipientIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$recipients | ForEach-Object { if ($_.ExternalDirectoryObjectId) { [void]$recipientIds.Add("$($_.ExternalDirectoryObjectId)") } }
$holding = @($recipients | Where-Object { @($_.EmailAddresses | Where-Object { $_ -match '^smtp:' -and (Test-OnDomain ($_ -replace '^smtp:', '')) }).Count })

# ------------------------------------------------------------ 1. snapshot
$snapFile = ''
if (-not $BlockersOnly) {
    $stage = if ($holding.Count) { 'BeforeRelease' } else { 'AfterRelease' }
    $tag = if ($holding.Count) { 'pre' } else { 'post' }
    $snapFile = ".\SourceRecipients-$tag-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
    $recipients | ForEach-Object {
        [pscustomobject]@{
            PrimarySmtpAddress   = "$($_.PrimarySmtpAddress)".ToLower()
            RecipientTypeDetails = $_.RecipientTypeDetails
            DisplayName          = $_.DisplayName
            Alias                = $_.Alias
            LegacyExchangeDN     = $_.LegacyExchangeDN
            EmailAddresses       = (@($_.EmailAddresses) -join ';')
            ObjectId             = $_.ExternalDirectoryObjectId
            SnapshotStage        = $stage
        }
    } | Export-Csv $snapFile -NoTypeInformation -WhatIf:$false
    if ($holding.Count) {
        Write-Host "Snapshot of $($recipients.Count) recipients ($($holding.Count) hold the domain) saved to $snapFile. Keep it: the destination needs it for X500 addresses." -ForegroundColor Green
    }
    else {
        Write-Host ''
        Write-Host '  !!! No recipient holds an address on the domain any more. !!!' -ForegroundColor Red
        Write-Host "  $snapFile was taken AFTER the addresses were removed. It can't be used for X500s or aliases." -ForegroundColor Red
        Write-Host '  Find the SourceRecipients-pre-*.csv file from before the release and use that with Switch-TargetAddresses.' -ForegroundColor Red
        Write-Host ''
    }
    if ($SnapshotOnly) { return }
}

# ------------------------------------------------------------ 2. blockers (always, WhatIf included)
$mgUsers = @(Get-MgUser -All -Property Id, UserPrincipalName, Mail, ProxyAddresses, OnPremisesSyncEnabled)
$apps = @(Get-MgApplication -All -Property Id, AppId, DisplayName, IdentifierUris, Web)

function Get-SourceBlockers([switch]$AfterChanges) {
    $found = [System.Collections.Generic.List[object]]::new()
    $handled = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    # Synced users: can't be renamed in the cloud
    foreach ($u in $mgUsers | Where-Object { $_.OnPremisesSyncEnabled -and (Test-OnDomain $_.UserPrincipalName) }) {
        $found.Add(@('Synced user', $u.UserPrincipalName, 'Blocker', 'Still synced from on-premises', 'Turn directory sync off (or change the UPN on-premises) first')); [void]$handled.Add($u.Id)
    }
    # Deleted users keep their sign-in name (prefixed with their ID) until purged or renamed
    $deleted = @(Get-MgDirectoryDeletedItemAsUser -All -Property Id, UserPrincipalName -ErrorAction SilentlyContinue |
        Where-Object { $_.UserPrincipalName -and (Test-OnDomain $_.UserPrincipalName) })
    foreach ($du in $deleted) {
        $found.Add(@('Deleted user', $du.UserPrincipalName, 'Blocker', 'In the recycle bin',
                "Run Clear-DeletedUserDomain.ps1, or purge with Remove-MgDirectoryDeletedItem -DirectoryObjectId $($du.Id)")); [void]$handled.Add($du.Id)
    }
    # App registrations: identifier URIs block removal; redirect URIs just break sign-in to the app
    foreach ($a in $apps) {
        foreach ($uri in @($a.IdentifierUris) | Where-Object { Test-HostOnDomain $_ }) {
            $new = Get-NewAppUri $uri
            $status = if ($FixApplications -and -not $AfterChanges) { 'Info' } else { 'Blocker' }
            $found.Add(@('Application', "$($a.DisplayName) ($($a.AppId))", $status, "Identifier URI $uri",
                    "Change it to $new (Entra admin centre > App registrations > Expose an API, or rerun with -FixApplications)"))
            [void]$handled.Add($a.Id)
        }
        foreach ($uri in @($a.Web.RedirectUris) | Where-Object { Test-HostOnDomain $_ }) {
            $found.Add(@('Application', "$($a.DisplayName) ($($a.AppId))", 'Warn', "Redirect URI $uri",
                    'Not a blocker, but the app breaks if that host stops resolving. Plan its new home'))
        }
    }
    # ProxyOnly: proxy addresses in Entra with no Exchange recipient to edit them through
    foreach ($u in $mgUsers | Where-Object { -not $recipientIds.Contains("$($_.Id)") }) {
        $hits = @(@($u.ProxyAddresses) | Where-Object { $_ -match '^smtp:' -and (Test-OnDomain ($_ -replace '^smtp:', '')) })
        if ($u.Mail -and (Test-OnDomain $u.Mail)) { $hits += "mail:$($u.Mail)" }
        if ($hits) {
            $found.Add(@('ProxyOnly', $u.UserPrincipalName, 'Blocker', "No mailbox, but Entra holds $($hits -join ', ')",
                    'Graph cannot edit proxyAddresses. Use the Microsoft 365 admin centre: Settings > Domains > the domain > Remove domain, and let the wizard move them'))
            [void]$handled.Add($u.Id)
        }
    }
    # Anything else the directory still links to the domain
    foreach ($d in $Domain) {
        $refs = @(Get-MgDomainNameReference -DomainId $d -All -ErrorAction SilentlyContinue)
        $left = foreach ($ref in $refs) {
            if ($handled.Contains("$($ref.Id)")) { continue }
            # Before changes, users and recipients this script will move aren't blockers
            if (-not $AfterChanges -and ($recipientIds.Contains("$($ref.Id)") -or ($mgUsers | Where-Object { $_.Id -eq $ref.Id -and (Test-OnDomain $_.UserPrincipalName) }))) { continue }
            $ref
        }
        foreach ($ref in $left) {
            $name = $ref.AdditionalProperties.userPrincipalName, $ref.AdditionalProperties.mail, $ref.AdditionalProperties.displayName | Where-Object { $_ } | Select-Object -First 1
            $found.Add(@('Reference', $d, 'Blocker', "$($ref.AdditionalProperties.'@odata.type') $name", 'Remove the domain from this object by hand, or use the admin centre Remove domain wizard'))
        }
        if (-not $left) { $found.Add(@('Reference', $d, 'Info', 'Nothing else in the directory blocks it', '')) }
    }
    $found
}
function Get-NewAppUri($uri) {
    $u = [uri]"$uri"
    "$($u.Scheme)://$initial$($u.PathAndQuery)".TrimEnd('/')
}

Write-Host ''
Write-Host 'Checking for blockers...' -ForegroundColor Cyan
foreach ($b in Get-SourceBlockers) { Write-Log $b[0] $b[1] $b[2] $b[3] $b[4] }
$hard = @($log | Where-Object { $_.Status -eq 'Blocker' -and $_.Area -in 'Synced user', 'Deleted user', 'Application' })

if ($BlockersOnly) {
    $file = ".\SourceDomainBlockers-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
    $log | Export-Csv $file -NoTypeInformation -WhatIf:$false
    $n = @($log | Where-Object Status -eq 'Blocker').Count
    Write-Host ''
    Write-Host "$n blocker(s). Report: $file" -ForegroundColor $(if ($n) { 'Yellow' } else { 'Green' })
    return
}
if ($hard -and -not $WhatIfPreference) {
    $file = ".\RemoveSourceDomain-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
    $log | Export-Csv $file -NoTypeInformation -WhatIf:$false
    Write-Host "$($hard.Count) hard blocker(s) above. Nothing has been changed. Fix them, then run again. Report: $file" -ForegroundColor Red
    return
}

# ------------------------------------------------------------ 3. applications
if ($FixApplications) {
    foreach ($a in $apps) {
        $uris = @($a.IdentifierUris)
        if (-not ($uris | Where-Object { Test-HostOnDomain $_ })) { continue }
        $newList = @($uris | ForEach-Object { if (Test-HostOnDomain $_) { Get-NewAppUri $_ } else { $_ } } | Select-Object -Unique)
        $label = "$($a.DisplayName) ($($a.AppId))"
        if ($PSCmdlet.ShouldProcess($label, "Identifier URIs -> $($newList -join ', ')")) {
            try { Update-MgApplication -ApplicationId $a.Id -IdentifierUris $newList; Write-Log 'Application' $label 'Done' "Identifier URIs now $($newList -join ', ')" }
            catch { Write-Log 'Application' $label 'Problem' $_.Exception.Message 'Change it in the Entra admin centre > App registrations > Expose an API' }
        }
        else { Write-Log 'Application' $label 'WhatIf' "Identifier URIs -> $($newList -join ', ')" }
    }
}

# ------------------------------------------------------------ 4. default domain
$default = ($domains | Where-Object IsDefault).Id
if ($Domain -contains $default) {
    if ($PSCmdlet.ShouldProcess($initial, 'Make the default domain')) {
        Update-MgDomain -DomainId $initial -IsDefault:$true
        Write-Log 'Domain' $initial 'Done' "Now the default (was $default)"
    }
    else { Write-Log 'Domain' $initial 'WhatIf' "Would become the default (is $default)" }
}

# ------------------------------------------------------------ 5. sign-in names
$users = @($mgUsers | Where-Object { Test-OnDomain $_.UserPrincipalName })
$taken = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$mgUsers | ForEach-Object { [void]$taken.Add($_.UserPrincipalName) }
foreach ($u in $users) {
    if ($u.OnPremisesSyncEnabled) { continue }   # already logged as a blocker
    $local = $u.UserPrincipalName.Split('@')[0]
    $new = "$local@$initial"; $n = 1
    while ($taken.Contains($new)) { $new = "$local$n@$initial"; $n++ }
    if ($PSCmdlet.ShouldProcess($u.UserPrincipalName, "Change sign-in to $new")) {
        try { Update-MgUser -UserId $u.Id -UserPrincipalName $new; [void]$taken.Add($new); Write-Log 'Sign-in' $u.UserPrincipalName 'Done' "-> $new" }
        catch { Write-Log 'Sign-in' $u.UserPrincipalName 'Problem' $_.Exception.Message }
    }
    else { Write-Log 'Sign-in' $u.UserPrincipalName 'WhatIf' "-> $new" }
}

# ------------------------------------------------------------ 6. addresses
$retry = [System.Collections.Generic.List[object]]::new()
function Invoke-AddressRemoval($Item) {
    # Returns 'Done', 'SyncLag' or throws. The WindowsLiveId hold shows up as a warning, not an error.
    $wv = $null
    & $Item.Set -Identity $Item.Id -EmailAddresses @{ Remove = $Item.Remove } -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable wv
    if (@($wv) -match 'WindowsLiveId') { return 'SyncLag' }
    'Done'
}
foreach ($r in $holding) {
    $all = @($r.EmailAddresses | Where-Object { $_ -match '^smtp:' })
    $onDomain = @($all | Where-Object { Test-OnDomain ($_ -replace '^smtp:', '') })
    $label = "$($r.PrimarySmtpAddress) ($($r.RecipientTypeDetails))"
    $type = "$($r.RecipientTypeDetails)"

    $set, $primaryParam = switch -Regex ($type) {
        '^GroupMailbox$'                   { 'Set-UnifiedGroup', 'PrimarySmtpAddress'; break }
        'Mailbox$'                         { 'Set-Mailbox', 'WindowsEmailAddress'; break }
        '^Mail(Universal|NonUniversal)'    { 'Set-DistributionGroup', 'PrimarySmtpAddress'; break }
        '^DynamicDistributionGroup$'       { 'Set-DynamicDistributionGroup', 'PrimarySmtpAddress'; break }
        '^(MailUser|GuestMailUser)$'       { 'Set-MailUser', 'PrimarySmtpAddress'; break }
        '^MailContact$'                    { 'Set-MailContact', 'PrimarySmtpAddress'; break }
        default                            { $null, $null }
    }
    if (-not $set) { Write-Log 'Address' $label 'Problem' 'Recipient type not handled' 'Change it by hand in the Exchange admin centre'; continue }

    $id = if ($r.ExternalDirectoryObjectId) { "$($r.ExternalDirectoryObjectId)" } else { "$($r.Identity)" }
    $item = $null
    try {
        if (Test-OnDomain $r.PrimarySmtpAddress) {
            $existing = $all | ForEach-Object { $_ -replace '^smtp:', '' } | Where-Object { $_ -like "*@$initial" } | Select-Object -First 1
            $newPrimary = if ($existing) { $existing } else { "$($r.Alias)@$initial" }
            if ($PSCmdlet.ShouldProcess($label, "Primary address -> $newPrimary")) {
                $p = @{ Identity = $id; ErrorAction = 'Stop' }; $p[$primaryParam] = $newPrimary
                & $set @p
                Write-Log 'Address' $label 'Done' "Primary -> $newPrimary"
            }
            else { Write-Log 'Address' $label 'WhatIf' "Primary -> $newPrimary" }
        }
        $item = [pscustomobject]@{ Label = $label; Set = $set; Id = $id; Remove = @($onDomain | ForEach-Object { 'smtp:' + ($_ -replace '^smtp:', '') }) }
        if ($PSCmdlet.ShouldProcess($label, "Remove $($item.Remove.Count) address(es) on $($Domain -join ', ')")) {
            if ((Invoke-AddressRemoval $item) -eq 'SyncLag') {
                $retry.Add($item)
                Write-Log 'Address' $label 'SyncLag' "Exchange still uses an address as WindowsLiveId. Will retry for $RetryMinutes min"
            }
            else { Write-Log 'Address' $label 'Done' "Removed $($item.Remove -join ', ')" }
        }
        else { Write-Log 'Address' $label 'WhatIf' "Would remove $($item.Remove -join ', ')" }
    }
    catch {
        if ($item -and "$($_.Exception.Message)" -match 'WindowsLiveId') { $retry.Add($item); Write-Log 'Address' $label 'SyncLag' 'WindowsLiveId hold. Will retry' }
        else { Write-Log 'Address' $label 'Problem' $_.Exception.Message }
    }
}

# Sign-in changes take a few minutes to reach Exchange. Retry until they have.
if ($retry.Count -and $RetryMinutes -gt 0) {
    $deadline = (Get-Date).AddMinutes($RetryMinutes)
    while ($retry.Count -and (Get-Date) -lt $deadline) {
        Write-Host "  Waiting 60s for sync, then retrying $($retry.Count) recipient(s)..." -ForegroundColor Magenta
        Start-Sleep -Seconds 60
        foreach ($item in @($retry)) {
            try {
                if ((Invoke-AddressRemoval $item) -eq 'Done') { [void]$retry.Remove($item); Write-Log 'Address' $item.Label 'Done' "Removed $($item.Remove -join ', ') after retry" }
            }
            catch { if ("$($_.Exception.Message)" -notmatch 'WindowsLiveId') { [void]$retry.Remove($item); Write-Log 'Address' $item.Label 'Problem' $_.Exception.Message } }
        }
    }
}
foreach ($item in $retry) {
    Write-Log 'Address' $item.Label 'Problem' "Still held as WindowsLiveId after $RetryMinutes min: $($item.Remove -join ', ')" 'Wait a few more minutes and run the script again (it is safe to re-run)'
}

# ------------------------------------------------------------ 7. remove
if ($RemoveDomain) {
    if (-not $WhatIfPreference) {
        Write-Host ''
        Write-Host 'Re-checking blockers after the changes...' -ForegroundColor Cyan
        $mgUsers = @(Get-MgUser -All -Property Id, UserPrincipalName, Mail, ProxyAddresses, OnPremisesSyncEnabled)
        $apps = @(Get-MgApplication -All -Property Id, AppId, DisplayName, IdentifierUris, Web)
        foreach ($b in Get-SourceBlockers -AfterChanges) { if ($b[2] -ne 'Info') { Write-Log $b[0] $b[1] $b[2] $b[3] $b[4] } }
    }
    $blockers = @($log | Where-Object { $_.Status -in 'Blocker', 'Problem' })
    if ($blockers -and -not $WhatIfPreference) {
        Write-Host "$($blockers.Count) blocker(s) or problem(s) above. Not removing. Fix them and run again with -RemoveDomain." -ForegroundColor Red
        if ($blockers | Where-Object Area -eq 'ProxyOnly') {
            Write-Host 'ProxyOnly users: finish in the Microsoft 365 admin centre, Settings > Domains > the domain > Remove domain.' -ForegroundColor Yellow
        }
    }
    else {
        if ($blockers -and $WhatIfPreference) { Write-Host "  On a real run, removal would be refused until the $($blockers.Count) blocker(s) above are cleared." -ForegroundColor Yellow }
        foreach ($d in $Domain) {
            if ($PSCmdlet.ShouldProcess($d, 'Disable DKIM and remove the domain from the tenant')) {
                try {
                    $dkim = Get-DkimSigningConfig -Identity $d -ErrorAction SilentlyContinue
                    if ($dkim) { Set-DkimSigningConfig -Identity $d -Enabled $false -ErrorAction SilentlyContinue | Out-Null }
                    Remove-MgDomain -DomainId $d
                    Write-Log 'Domain' $d 'Done' 'Removal started. It can take a while to finish; the destination can verify it once it has'
                }
                catch { Write-Log 'Domain' $d 'Problem' $_.Exception.Message }
            }
            else { Write-Log 'Domain' $d 'WhatIf' 'Would disable DKIM and remove the domain' }
        }
    }
}

$file = ".\RemoveSourceDomain-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
$log | Export-Csv $file -NoTypeInformation -WhatIf:$false
Write-Host ''
$log | Group-Object Status | ForEach-Object { Write-Host ("{0,-8} {1}" -f $_.Name, $_.Count) }
Write-Host "Report: $file   Snapshot: $snapFile" -ForegroundColor Cyan
if ($RemoveDomain -and ($log | Where-Object { $_.Area -eq 'Domain' -and $_.Status -eq 'Done' -and $_.Object -ne $initial })) {
    Write-Host "Next: in the destination, run .\Add-TargetDomain.ps1 -TenantId <destination-guid> -Domain $($Domain -join ', ') -Verify" -ForegroundColor Green
}
else {
    Write-Host "Next: clear anything above, then run with -RemoveDomain. After that, Add-TargetDomain -Verify in the destination." -ForegroundColor Cyan
}
