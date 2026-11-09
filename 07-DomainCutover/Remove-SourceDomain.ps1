<#
.SYNOPSIS
    Releases domains from the source tenant: snapshots every recipient, moves sign-in names and addresses off the
    domains, lists anything still holding on, and removes the domains.

.DESCRIPTION
    TL;DR: A domain can only live in one tenant. Before the destination can verify it, the source has to let go.
    This script:
      1. Snapshots every recipient (addresses and LegacyExchangeDN) to SourceRecipients-<date>.csv. The destination
         needs that file to add X500 addresses, so replies to old mail keep working.
      2. Moves the default domain to onmicrosoft.com if one of the domains is the default.
      3. Changes every sign-in name on the domains to <name>@<tenant>.onmicrosoft.com.
      4. Changes every recipient's primary address off the domains and removes their addresses on them.
      5. Lists what still references the domains (Graph), including deleted users still in the recycle bin.
      6. With -RemoveDomain, disables DKIM and removes the domains.
    Steps 2 to 6 support -WhatIf. -SnapshotOnly does step 1 and stops.

    This is the point of no return for mail. From the moment the domain leaves the source until MX points at the
    destination, mail to it can bounce. Lower DNS TTLs days before, and do this and Add-TargetDomain back to back.

.PARAMETER Domain
    Domains to release, e.g. fabrikam.com. Never the source's onmicrosoft.com domain.

.PARAMETER TenantId
    Source tenant ID (GUID). Required: the script refuses to touch any other tenant.

.PARAMETER SnapshotOnly
    Write the recipient snapshot and stop. Run this first, today, so you have it whatever happens.

.PARAMETER RemoveDomain
    After moving everything off, remove the domains from the tenant.

.EXAMPLE
    .\Remove-SourceDomain.ps1 -TenantId <source-guid> -Domain fabrikam.com -SnapshotOnly

.EXAMPLE
    .\Remove-SourceDomain.ps1 -TenantId <source-guid> -Domain fabrikam.com, fabrikam.co.uk -WhatIf

.EXAMPLE
    .\Remove-SourceDomain.ps1 -TenantId <source-guid> -Domain fabrikam.com, fabrikam.co.uk -RemoveDomain

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.0.1
    Date    : 2026-11-09
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: ExchangeOnlineManagement 3.x, Microsoft.Graph.Authentication, Microsoft.Graph.Users,
              Microsoft.Graph.Identity.DirectoryManagement. Global Administrator in the source.
    Directory sync must be off for the source: synced users' sign-in names and addresses can't be changed in the cloud.
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)][string[]]$Domain,
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F-]{36}$')][string]$TenantId,
    [switch]$SnapshotOnly,
    [switch]$RemoveDomain
)

$ErrorActionPreference = 'Stop'
$Domain = @($Domain | ForEach-Object { $_.Trim().ToLower() })
if ($Domain | Where-Object { $_ -like '*.onmicrosoft.com' }) { throw "Don't release an onmicrosoft.com domain." }

# ------------------------------------------------------------ connect, both to the same tenant
if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object State -eq 'Connected')) { Connect-ExchangeOnline -ShowBanner:$false }
$exoTenant = (Get-ConnectionInformation | Where-Object State -eq 'Connected' | Select-Object -First 1).TenantID
if ($exoTenant -ne $TenantId) { throw "Exchange Online is signed in to $exoTenant, not $TenantId. Disconnect-ExchangeOnline and sign in to the source." }
Connect-MgGraph -TenantId $TenantId -NoWelcome -Scopes 'User.ReadWrite.All', 'Domain.ReadWrite.All', 'Directory.ReadWrite.All'
if ((Get-MgContext).TenantId -ne $TenantId) { throw "Graph signed in to $((Get-MgContext).TenantId), not $TenantId." }

$domains = @(Get-MgDomain -All)
$initial = ($domains | Where-Object IsInitial).Id
foreach ($d in $Domain) { if ($domains.Id -notcontains $d) { throw "$d isn't a domain in this tenant." } }
Write-Host "Source tenant confirmed. Fallback domain: $initial" -ForegroundColor Cyan

function Test-OnDomain($address) { $Domain -contains ("$address".Split('@')[-1].ToLower()) }
$log = [System.Collections.Generic.List[object]]::new()
function Write-Log($Area, $Object, $Status, $Detail) {
    $log.Add([pscustomobject]@{ Area = $Area; Object = $Object; Status = $Status; Detail = $Detail })
    $c = @{ Done = 'Green'; WhatIf = 'Cyan'; Problem = 'Red'; Info = 'Gray'; Blocker = 'Yellow' }[$Status]
    Write-Host ("  [{0}] {1}: {2} {3}" -f $Status, $Area, $Object, $Detail) -ForegroundColor $c
}

# ------------------------------------------------------------ 1. snapshot
$recipients = @(Get-EXORecipient -ResultSize Unlimited -Properties LegacyExchangeDN, EmailAddresses, ExternalDirectoryObjectId)
$snapFile = ".\SourceRecipients-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
$recipients | ForEach-Object {
    [pscustomobject]@{
        PrimarySmtpAddress   = "$($_.PrimarySmtpAddress)".ToLower()
        RecipientTypeDetails = $_.RecipientTypeDetails
        DisplayName          = $_.DisplayName
        Alias                = $_.Alias
        LegacyExchangeDN     = $_.LegacyExchangeDN
        EmailAddresses       = (@($_.EmailAddresses) -join ';')
        ObjectId             = $_.ExternalDirectoryObjectId
    }
} | Export-Csv $snapFile -NoTypeInformation -WhatIf:$false
Write-Host "Snapshot of $($recipients.Count) recipients saved to $snapFile. Keep it: the destination needs it for X500 addresses." -ForegroundColor Green
if ($SnapshotOnly) { return }

# ------------------------------------------------------------ 2. default domain
$default = ($domains | Where-Object IsDefault).Id
if ($Domain -contains $default) {
    if ($PSCmdlet.ShouldProcess($initial, 'Make the default domain')) {
        Update-MgDomain -DomainId $initial -IsDefault:$true
        Write-Log 'Domain' $initial 'Done' "Now the default (was $default)"
    }
    else { Write-Log 'Domain' $initial 'WhatIf' "Would become the default (is $default)" }
}

# ------------------------------------------------------------ 3. sign-in names
$users = @(Get-MgUser -All -Property Id, UserPrincipalName, OnPremisesSyncEnabled | Where-Object { Test-OnDomain $_.UserPrincipalName })
$taken = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
Get-MgUser -All -Property UserPrincipalName | ForEach-Object { [void]$taken.Add($_.UserPrincipalName) }
foreach ($u in $users) {
    if ($u.OnPremisesSyncEnabled) { Write-Log 'Sign-in' $u.UserPrincipalName 'Problem' 'Still synced from on-premises. Turn sync off first'; continue }
    $local = $u.UserPrincipalName.Split('@')[0]
    $new = "$local@$initial"; $n = 1
    while ($taken.Contains($new)) { $new = "$local$n@$initial"; $n++ }
    if ($PSCmdlet.ShouldProcess($u.UserPrincipalName, "Change sign-in to $new")) {
        try { Update-MgUser -UserId $u.Id -UserPrincipalName $new; [void]$taken.Add($new); Write-Log 'Sign-in' $u.UserPrincipalName 'Done' "-> $new" }
        catch { Write-Log 'Sign-in' $u.UserPrincipalName 'Problem' $_.Exception.Message }
    }
    else { Write-Log 'Sign-in' $u.UserPrincipalName 'WhatIf' "-> $new" }
}

# ------------------------------------------------------------ 4. addresses
foreach ($r in $recipients) {
    $all = @($r.EmailAddresses | Where-Object { $_ -match '^smtp:' })
    $onDomain = @($all | Where-Object { Test-OnDomain ($_ -replace '^smtp:', '') })
    if (-not $onDomain) { continue }
    $label = "$($r.PrimarySmtpAddress) ($($r.RecipientTypeDetails))"
    $type = "$($r.RecipientTypeDetails)"

    $set, $primaryParam = switch -Regex ($type) {
        'Mailbox$' { if ($type -eq 'GroupMailbox') { 'Set-UnifiedGroup', 'PrimarySmtpAddress' } else { 'Set-Mailbox', 'WindowsEmailAddress' } }
        '^Mail(Universal|NonUniversal)'    { 'Set-DistributionGroup', 'PrimarySmtpAddress' }
        '^DynamicDistributionGroup$'       { 'Set-DynamicDistributionGroup', 'PrimarySmtpAddress' }
        '^(MailUser|GuestMailUser)$'       { 'Set-MailUser', 'PrimarySmtpAddress' }
        '^MailContact$'                    { 'Set-MailContact', 'PrimarySmtpAddress' }
        default                            { $null, $null }
    }
    if (-not $set) { Write-Log 'Address' $label 'Problem' 'Recipient type not handled. Change it by hand'; continue }

    $id = if ($r.ExternalDirectoryObjectId) { $r.ExternalDirectoryObjectId } else { "$($r.Identity)" }
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
        $remove = @($onDomain | ForEach-Object { 'smtp:' + ($_ -replace '^smtp:', '') })
        if ($PSCmdlet.ShouldProcess($label, "Remove $($remove.Count) address(es) on $($Domain -join ', ')")) {
            & $set -Identity $id -EmailAddresses @{ Remove = $remove } -ErrorAction Stop
            Write-Log 'Address' $label 'Done' "Removed $($remove -join ', ')"
        }
        else { Write-Log 'Address' $label 'WhatIf' "Would remove $($remove -join ', ')" }
    }
    catch { Write-Log 'Address' $label 'Problem' $_.Exception.Message }
}

# ------------------------------------------------------------ 5. what still holds on
foreach ($d in $Domain) {
    $refs = @(Get-MgDomainNameReference -DomainId $d -All -ErrorAction SilentlyContinue)
    if ($refs -and -not $WhatIfPreference) {
        foreach ($ref in $refs) {
            $name = $ref.AdditionalProperties.userPrincipalName, $ref.AdditionalProperties.mail, $ref.AdditionalProperties.displayName | Where-Object { $_ } | Select-Object -First 1
            Write-Log 'Reference' $d 'Blocker' "$($ref.AdditionalProperties.'@odata.type') $name"
        }
    }
    elseif (-not $WhatIfPreference) { Write-Log 'Reference' $d 'Info' 'Nothing in the directory references it' }
}
# Deleted users keep their sign-in name (prefixed with their ID) and block removal until purged or renamed.
$deleted = @(Get-MgDirectoryDeletedItemAsUser -All -Property Id, UserPrincipalName -ErrorAction SilentlyContinue |
    Where-Object { $_.UserPrincipalName -and (Test-OnDomain $_.UserPrincipalName) })
foreach ($du in $deleted) {
    Write-Log 'Deleted user' $du.UserPrincipalName 'Blocker' "In the recycle bin. Restore and rename, or purge with Remove-MgDirectoryDeletedItem -DirectoryObjectId $($du.Id)"
}

# ------------------------------------------------------------ 6. remove
if ($RemoveDomain) {
    $blockers = @($log | Where-Object { $_.Status -in 'Blocker', 'Problem' })
    if ($blockers -and -not $WhatIfPreference) {
        Write-Host "$($blockers.Count) blocker(s) or problem(s) above. Not removing. Fix them and run again with -RemoveDomain." -ForegroundColor Red
    }
    else {
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
