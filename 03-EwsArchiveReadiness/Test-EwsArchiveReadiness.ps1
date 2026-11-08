<#
.SYNOPSIS
    Checks, and optionally fixes, everything that stops a migration tool reading or writing In-Place Archives over EWS.

.DESCRIPTION
    TL;DR: Run it in the source AND the destination tenant with your migration tool's app ID. It checks four things:
    the tenant EWS switch and app ID allow list, any old user-agent allow list left behind, mailboxes with EWS
    switched off on their own, and whether the app has EWS permission. Add -CheckArchives to see which archives
    are actually ready. Add -Fix to correct what it finds. Everything supports -WhatIf.

    Why it exists: many migration tools copy the main mailbox through Microsoft Graph but still use EWS for
    In-Place Archives. Since October 2026 EWS is off by default and only apps on the tenant's EwsAllowedAppIDs
    list get in. Get any one of these checks wrong and the main mailbox copies fine while every archive fails
    with 'Unauthorized'.

.PARAMETER AppId
    The application (client) ID your migration tool uses for EWS.

.PARAMETER TenantId
    Optional tenant ID (GUID). Graph signs in to this tenant, and the script stops (and disconnects) if Exchange
    is signed in to any other. Saves you fixing the wrong side of a migration.

.PARAMETER Fix
    Correct what the checks find: enable EWS for the tenant, add AppId to EwsAllowedAppIDs (keeping what's already
    there), clear the old user-agent policy and allow list, and re-enable EWS on blocked mailboxes.

.PARAMETER Mailbox
    Limit the per-mailbox checks to these addresses. All mailboxes if omitted.

.PARAMETER CheckArchives
    Also check that every active archive can be read (Get-EXOMailboxStatistics -Archive). A brand-new archive
    can show ArchiveStatus Active for a while before it's ready.

.PARAMETER SkipConsentCheck
    Skip the Microsoft Graph check for the full_access_as_app permission.

.EXAMPLE
    .\Test-EwsArchiveReadiness.ps1 -AppId 00000000-0000-0000-0000-000000000000

.EXAMPLE
    .\Test-EwsArchiveReadiness.ps1 -AppId 00000000-0000-0000-0000-000000000000 -Fix -WhatIf

.EXAMPLE
    .\Test-EwsArchiveReadiness.ps1 -AppId 00000000-0000-0000-0000-000000000000 -TenantId 11111111-1111-1111-1111-111111111111 -Fix -CheckArchives

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.0.2
    Date    : 2026-11-08
    Changelog:
      1.0.2  Clears the legacy allow and block lists in separate calls (Exchange rejects both at once) and only
             reports Fixed when Exchange accepts the change. Reads list values whether they come back as a string or a collection.
      1.0.1  -TenantId: Graph signs in to the named tenant, and an Exchange session for any other tenant is refused.
      1.0.0  First release.
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: ExchangeOnlineManagement 3.x. For the consent check: Microsoft.Graph.Authentication and
              Microsoft.Graph.Applications. Read-only without -Fix. -Fix needs Exchange Administrator; the consent
              check needs Application.Read.All.
    EWS in Exchange Online is retired for good on 1 April 2027. Plan archive copies before then.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F-]{36}$')][string]$AppId,
    [ValidatePattern('^[0-9a-fA-F-]{36}$')][string]$TenantId,
    [switch]$Fix,
    [string[]]$Mailbox,
    [switch]$CheckArchives,
    [switch]$SkipConsentCheck
)

$ErrorActionPreference = 'Stop'
$findings = [System.Collections.Generic.List[object]]::new()
function Add-Finding($Area, $Target, $Status, $Detail) {
    $findings.Add([pscustomobject]@{ Area = $Area; Target = $Target; Status = $Status; Detail = $Detail })
    $colour = @{ OK = 'Green'; Fixed = 'Cyan'; Problem = 'Red'; Warning = 'Yellow'; Info = 'Gray' }[$Status]
    Write-Host ("[{0}] {1}: {2} {3}" -f $Status, $Area, $Target, $Detail) -ForegroundColor $colour
}

if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object State -eq 'Connected')) {
    Connect-ExchangeOnline -ShowBanner:$false
}
$conn = Get-ConnectionInformation | Where-Object State -eq 'Connected' | Select-Object -First 1
if ($TenantId -and $conn.TenantID -ne $TenantId) {
    $wrong = $conn.TenantID
    Disconnect-ExchangeOnline -Confirm:$false
    throw "Exchange signed in to tenant $wrong, not $TenantId. Run the script again and sign in with an admin of the right tenant."
}
$tenant = (Get-AcceptedDomain | Where-Object Default).DomainName
Write-Host "Tenant: $tenant" -ForegroundColor Cyan

# ------------------------------------------------------------------ 1. tenant switch and app ID allow list
$org = Get-OrganizationConfig -RetrieveEwsOperationAccessPolicy
# Exchange can return the list as one comma-separated string or as a collection. Handle both.
function Get-Values($v) { @(@($v) | ForEach-Object { "$_" -split '[,;\s]+' } | ForEach-Object { $_.Trim('{}', ' ') } | Where-Object { $_ }) }
$allowed = Get-Values $org.EwsAllowedAppIDs

if ($org.EwsEnabled -eq $true) { Add-Finding 'Tenant' 'EwsEnabled' 'OK' 'True' }
elseif ($Fix -and $PSCmdlet.ShouldProcess($tenant, 'Enable EWS')) {
    Set-OrganizationConfig -EwsEnabled $true -ErrorAction Stop
    Add-Finding 'Tenant' 'EwsEnabled' 'Fixed' 'Set to True'
}
else { Add-Finding 'Tenant' 'EwsEnabled' 'Problem' "Is '$($org.EwsEnabled)'. EWS is off for the whole tenant" }

if ($allowed -contains $AppId) { Add-Finding 'Tenant' 'EwsAllowedAppIDs' 'OK' "$AppId is on the list" }
elseif ($Fix -and $PSCmdlet.ShouldProcess($tenant, "Add $AppId to EwsAllowedAppIDs")) {
    # The parameter replaces the whole list, so keep what's already there.
    Set-OrganizationConfig -EwsAllowedAppIDs ((@($allowed) + $AppId | Select-Object -Unique) -join ',') -ErrorAction Stop
    Add-Finding 'Tenant' 'EwsAllowedAppIDs' 'Fixed' "Added $AppId (kept $($allowed.Count) existing)"
}
else { Add-Finding 'Tenant' 'EwsAllowedAppIDs' 'Problem' "$AppId isn't on the list. Current: $($allowed -join ', ')" }

# ------------------------------------------------------------------ 2. old user-agent policy left behind
$legacyPolicy = "$($org.EwsApplicationAccessPolicy)".Trim()
$legacyAllow  = Get-Values $org.EwsAllowList
$legacyBlock  = Get-Values $org.EwsBlockList
if (-not ($legacyPolicy -or $legacyAllow -or $legacyBlock)) {
    Add-Finding 'Tenant' 'Legacy policy' 'OK' 'No user-agent policy, allow list or block list'
}
elseif ($Fix -and $PSCmdlet.ShouldProcess($tenant, 'Clear the legacy EWS user-agent policy and lists')) {
    $was = "policy '$legacyPolicy', allow '$($legacyAllow -join ';')', block '$($legacyBlock -join ';')'"
    try {
        # Exchange refuses EwsAllowList and EwsBlockList in the same call, so clear them one at a time.
        if ($legacyAllow)  { Set-OrganizationConfig -EwsAllowList $null -ErrorAction Stop }
        if ($legacyBlock)  { Set-OrganizationConfig -EwsBlockList $null -ErrorAction Stop }
        if ($legacyPolicy) { Set-OrganizationConfig -EwsApplicationAccessPolicy $null -ErrorAction Stop }
        Add-Finding 'Tenant' 'Legacy policy' 'Fixed' "Cleared. Was $was"
    }
    catch { Add-Finding 'Tenant' 'Legacy policy' 'Problem' "Couldn't clear ($was): $($_.Exception.Message)" }
}
else {
    Add-Finding 'Tenant' 'Legacy policy' 'Warning' "Policy '$legacyPolicy', allow '$($legacyAllow -join ';')', block '$($legacyBlock -join ';')'. These match user agents, not app IDs, and can still block your tool"
}

# ------------------------------------------------------------------ 3. mailboxes with EWS off on their own
$cas = if ($Mailbox) { $Mailbox | ForEach-Object { Get-CASMailbox -Identity $_ } } else { Get-CASMailbox -ResultSize Unlimited }
$blocked = @($cas | Where-Object { $_.EwsEnabled -eq $false -or "$($_.EwsApplicationAccessPolicy)".Trim() -or (Get-Values $_.EwsAllowList) -or (Get-Values $_.EwsBlockList) })
if (-not $blocked) { Add-Finding 'Mailboxes' "$(@($cas).Count) checked" 'OK' 'None block EWS on their own' }
foreach ($c in $blocked) {
    $why = "EwsEnabled '$($c.EwsEnabled)', policy '$($c.EwsApplicationAccessPolicy)', allow '$($c.EwsAllowList -join ';')', block '$($c.EwsBlockList -join ';')'"
    if ($Fix -and $PSCmdlet.ShouldProcess($c.PrimarySmtpAddress, 'Re-enable EWS and clear mailbox-level lists')) {
        try {
            Set-CASMailbox -Identity $c.Identity -EwsEnabled $true -ErrorAction Stop
            if (Get-Values $c.EwsAllowList) { Set-CASMailbox -Identity $c.Identity -EwsAllowList $null -ErrorAction Stop }
            if (Get-Values $c.EwsBlockList) { Set-CASMailbox -Identity $c.Identity -EwsBlockList $null -ErrorAction Stop }
            if ("$($c.EwsApplicationAccessPolicy)".Trim()) { Set-CASMailbox -Identity $c.Identity -EwsApplicationAccessPolicy $null -ErrorAction Stop }
            Add-Finding 'Mailbox' $c.PrimarySmtpAddress 'Fixed' "Was $why"
        }
        catch { Add-Finding 'Mailbox' $c.PrimarySmtpAddress 'Problem' "Couldn't fix ($why): $($_.Exception.Message)" }
    }
    else { Add-Finding 'Mailbox' $c.PrimarySmtpAddress 'Problem' $why }
}

# ------------------------------------------------------------------ 4. does the app have EWS permission?
if (-not $SkipConsentCheck) {
    try {
        $ctx = Get-MgContext
        if ($TenantId -and $ctx -and $ctx.TenantId -ne $TenantId) { Disconnect-MgGraph | Out-Null; $ctx = $null }
        if (-not $ctx -or $ctx.Scopes -notcontains 'Application.Read.All') {
            if ($TenantId) { Connect-MgGraph -TenantId $TenantId -Scopes 'Application.Read.All' -NoWelcome }
            else { Connect-MgGraph -Scopes 'Application.Read.All' -NoWelcome }
        }
        $sp  = Get-MgServicePrincipal -Filter "appId eq '$AppId'"
        $exo = Get-MgServicePrincipal -Filter "appId eq '00000002-0000-0ff1-ce00-000000000000'"   # Office 365 Exchange Online
        if (-not $sp) { Add-Finding 'Consent' $AppId 'Problem' 'No service principal. The app was never consented in this tenant' }
        else {
            $roles = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id | Where-Object ResourceId -eq $exo.Id |
                ForEach-Object { $id = $_.AppRoleId; ($exo.AppRoles | Where-Object Id -eq $id).Value }
            if ($roles -contains 'full_access_as_app') { Add-Finding 'Consent' $sp.DisplayName 'OK' 'Has full_access_as_app' }
            else { Add-Finding 'Consent' $sp.DisplayName 'Problem' "No full_access_as_app (has: $($roles -join ', ')). Re-consent the app as a Global Admin" }
        }
    }
    catch { Add-Finding 'Consent' $AppId 'Warning' "Couldn't check: $($_.Exception.Message)" }
}

# ------------------------------------------------------------------ 5. are the archives actually ready?
if ($CheckArchives) {
    $mbx = if ($Mailbox) { $Mailbox | ForEach-Object { Get-EXOMailbox -Identity $_ -Properties ArchiveStatus } }
           else { Get-EXOMailbox -ResultSize Unlimited -Properties ArchiveStatus }
    $active = @($mbx | Where-Object ArchiveStatus -eq 'Active')
    $notReady = @($active | Where-Object { -not (Get-EXOMailboxStatistics -Identity $_.ExternalDirectoryObjectId -Archive -ErrorAction SilentlyContinue) })
    Add-Finding 'Archives' "$($active.Count) active" $(if ($notReady) { 'Warning' } else { 'OK' }) "$($active.Count - $notReady.Count) ready, $($notReady.Count) not ready yet"
    foreach ($n in $notReady) {
        Add-Finding 'Archive' $n.PrimarySmtpAddress 'Warning' 'Active but not readable yet. New archives can take hours. Re-check before copying'
    }
}

# ------------------------------------------------------------------ report
$file = ".\EwsArchiveReadiness-$($tenant)-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
$findings | Export-Csv $file -NoTypeInformation
$problems = @($findings | Where-Object Status -eq 'Problem').Count
Write-Host ''
Write-Host "Report: $file" -ForegroundColor Cyan
if ($problems) { Write-Host "$problems problem(s) found. Re-run with -Fix (try -WhatIf first)." -ForegroundColor Red }
elseif (@($findings | Where-Object Status -eq 'Fixed').Count) { Write-Host 'Fixes applied. Allow 15 to 30 minutes, then re-run without -Fix to confirm.' -ForegroundColor Cyan }
else { Write-Host 'No blockers found. Run it in the other tenant too.' -ForegroundColor Green }
