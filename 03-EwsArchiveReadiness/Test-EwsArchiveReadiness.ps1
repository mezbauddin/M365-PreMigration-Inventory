<#
.SYNOPSIS
    Checks, and optionally fixes, everything that stops a migration tool reading archives from the source tenant or
    writing them to the destination tenant over EWS. Does both tenants in one run.

.DESCRIPTION
    TL;DR: Give it your migration tool's app ID and both tenant IDs. It asks you to sign in to the source, checks
    it, then asks you to sign in to the destination and checks that. For each tenant it looks at the EWS switch, the
    app ID allow list, old user-agent rules left behind, mailboxes with EWS switched off on their own, the app's
    full_access_as_app permission and, with -CheckArchives, whether each archive is actually ready. -Fix corrects
    what it finds. Everything supports -WhatIf.

    Why it exists: many migration tools copy the main mailbox through Microsoft Graph but still use EWS for
    In-Place Archives. Since October 2026, once EWS is on, only apps on each tenant's EwsAllowedAppIDs
    list get in. Get any layer wrong, in either tenant, and the main mailbox copies fine while every archive fails
    with 'Unauthorized'.

.PARAMETER AppId
    The application (client) ID your migration tool uses for EWS.

.PARAMETER SourceTenantId
    Tenant ID (GUID) of the source. The tool has to READ archives here.

.PARAMETER DestinationTenantId
    Tenant ID (GUID) of the destination. The tool has to WRITE archives here.

.PARAMETER SourceAdmin
    Optional admin UPN for the source. Saves Windows from signing you in with a cached account from the other tenant.

.PARAMETER DestinationAdmin
    Optional admin UPN for the destination.

.PARAMETER Only
    Check just one side: Source or Destination. Both if omitted.

.PARAMETER Fix
    Correct what the checks find: enable EWS for the tenant, add AppId to EwsAllowedAppIDs (keeping what's already
    there), clear the old user-agent policy and its lists, and re-enable EWS on blocked mailboxes.

.PARAMETER CheckArchives
    Also check that every active archive can be read. A brand-new archive can show ArchiveStatus Active for hours
    before it's usable.

.PARAMETER SkipConsentCheck
    Skip the Microsoft Graph check for the full_access_as_app permission.

.EXAMPLE
    .\Test-EwsArchiveReadiness.ps1 -AppId <app-id> -SourceTenantId <source-guid> -DestinationTenantId <dest-guid>

.EXAMPLE
    .\Test-EwsArchiveReadiness.ps1 -AppId <app-id> -SourceTenantId <source-guid> -DestinationTenantId <dest-guid> -Fix -WhatIf

.EXAMPLE
    .\Test-EwsArchiveReadiness.ps1 -AppId <app-id> -SourceTenantId <source-guid> -DestinationTenantId <dest-guid> `
        -SourceAdmin admin@fabrikam.onmicrosoft.com -DestinationAdmin admin@contoso.onmicrosoft.com -Fix -CheckArchives

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 2.0.2
    Date    : 2026-11-09
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: ExchangeOnlineManagement 3.x (3.7.2 or later recommended, for -DisableWAM). For the consent check:
              Microsoft.Graph.Authentication and Microsoft.Graph.Applications. Read-only without -Fix. -Fix needs
              Exchange Administrator in each tenant; the consent check needs Application.Read.All.
    Changelog:
      2.0.2  The findings CSV is written under -WhatIf too (Export-Csv -WhatIf:$false), so a dry run leaves a report.
      2.0.1  Missing full_access_as_app is a Warning (delegated tools don't need it). Paged role lookup. Timing and wording per Microsoft's guidance.
      2.0.0  Checks source and destination in one run, with a sign-in for each and a guard against the wrong tenant.
      1.0.4  Clears legacy rules by setting the policy to null, which removes both lists.
      1.0.3  An active legacy EnforceAllowList is reported as a Problem.
      1.0.2  Only reports Fixed when Exchange accepts the change. Reads list values as a string or a collection.
      1.0.1  -TenantId guard.
      1.0.0  First release.
    EWS in Exchange Online is retired for good on 1 April 2027. Plan archive copies before then.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F-]{36}$')][string]$AppId,
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F-]{36}$')][string]$SourceTenantId,
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F-]{36}$')][string]$DestinationTenantId,
    [string]$SourceAdmin,
    [string]$DestinationAdmin,
    [ValidateSet('Source', 'Destination')][string]$Only,
    [switch]$Fix,
    [switch]$CheckArchives,
    [switch]$SkipConsentCheck
)

$ErrorActionPreference = 'Stop'
if ($SourceTenantId -eq $DestinationTenantId) { throw 'SourceTenantId and DestinationTenantId are the same tenant.' }
$Cmdlet   = $PSCmdlet
$findings = [System.Collections.Generic.List[object]]::new()
$canDisableWam = (Get-Command Connect-ExchangeOnline).Parameters.ContainsKey('DisableWAM')

function Add-Finding($Area, $Target, $Status, $Detail) {
    $findings.Add([pscustomobject]@{ Side = $script:Side; Tenant = $script:TenantName; Area = $Area; Target = $Target; Status = $Status; Detail = $Detail })
    $colour = @{ OK = 'Green'; Fixed = 'Cyan'; Problem = 'Red'; Warning = 'Yellow' }[$Status]
    Write-Host ("  [{0}] {1}: {2} {3}" -f $Status, $Area, $Target, $Detail) -ForegroundColor $colour
}

# Exchange can return a list as one comma-separated string or as a collection. Handle both.
function Get-Values($v) { @(@($v) | ForEach-Object { "$_" -split '[,;\s]+' } | ForEach-Object { $_.Trim([char[]]'{} ') } | Where-Object { $_ }) }

function Disconnect-All {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    if (Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
}

function Connect-Side($Side, $TenantId, $Admin) {
    Disconnect-All
    for ($attempt = 1; $attempt -le 2; $attempt++) {
        Write-Host ''
        Write-Host "Sign in to the $($Side.ToUpper()) tenant ($TenantId) with an Exchange admin." -ForegroundColor Magenta
        $p = @{ ShowBanner = $false }
        if ($Admin) { $p.UserPrincipalName = $Admin }
        # Second go: skip Windows' cached account so you get a fresh account picker.
        if ($attempt -eq 2 -and $canDisableWam) { $p.DisableWAM = $true }
        Connect-ExchangeOnline @p
        $conn = Get-ConnectionInformation | Where-Object State -eq 'Connected' | Select-Object -First 1
        if ($conn.TenantID -eq $TenantId) { return $true }
        Write-Warning "That sign-in went to tenant $($conn.TenantID), not the $Side tenant."
        Disconnect-ExchangeOnline -Confirm:$false | Out-Null
        if ($attempt -eq 1) { Write-Host 'Trying again. Pick the right account this time (or pass -SourceAdmin / -DestinationAdmin).' -ForegroundColor Yellow }
    }
    return $false
}

function Test-Side($Side, $TenantId, $Admin) {
    $script:Side = $Side; $script:TenantName = $TenantId
    Write-Host ''
    Write-Host "==================== $($Side.ToUpper()) ====================" -ForegroundColor Cyan
    if (-not (Connect-Side $Side $TenantId $Admin)) {
        Add-Finding 'Sign-in' $TenantId 'Problem' "Couldn't sign in to the $Side tenant. Nothing was checked or changed here"
        return
    }
    $script:TenantName = (Get-AcceptedDomain | Where-Object Default).DomainName
    Write-Host "Connected to $($script:TenantName)" -ForegroundColor Cyan
    $where = "$Side ($($script:TenantName))"

    # ---------------------------------------------------------- 1. tenant switch and app ID allow list
    $org = Get-OrganizationConfig -RetrieveEwsOperationAccessPolicy
    $allowed = Get-Values $org.EwsAllowedAppIDs

    if ($org.EwsEnabled -eq $true) { Add-Finding 'Tenant' 'EwsEnabled' 'OK' 'True' }
    elseif ($Fix -and $Cmdlet.ShouldProcess($where, 'Enable EWS')) {
        Set-OrganizationConfig -EwsEnabled $true -ErrorAction Stop
        Add-Finding 'Tenant' 'EwsEnabled' 'Fixed' 'Set to True'
    }
    else { Add-Finding 'Tenant' 'EwsEnabled' 'Problem' "Is '$($org.EwsEnabled)'. EWS is off for the whole tenant" }

    if ($allowed -contains $AppId) { Add-Finding 'Tenant' 'EwsAllowedAppIDs' 'OK' "$AppId is on the list" }
    elseif ($Fix -and $Cmdlet.ShouldProcess($where, "Add $AppId to EwsAllowedAppIDs")) {
        # The parameter replaces the whole list, so keep what's already there.
        Set-OrganizationConfig -EwsAllowedAppIDs ((@($allowed) + $AppId | Select-Object -Unique) -join ',') -ErrorAction Stop
        Add-Finding 'Tenant' 'EwsAllowedAppIDs' 'Fixed' "Added $AppId (kept $($allowed.Count) existing)"
    }
    else { Add-Finding 'Tenant' 'EwsAllowedAppIDs' 'Problem' "$AppId isn't on the list. Current: '$($allowed -join ', ')'" }

    # ---------------------------------------------------------- 2. old user-agent policy left behind
    $legacyPolicy = "$($org.EwsApplicationAccessPolicy)".Trim()
    $legacyAllow  = Get-Values $org.EwsAllowList
    $legacyBlock  = Get-Values $org.EwsBlockList
    $was = "policy '$legacyPolicy', allow '$($legacyAllow -join ';')', block '$($legacyBlock -join ';')'"
    if (-not ($legacyPolicy -or $legacyAllow -or $legacyBlock)) {
        Add-Finding 'Tenant' 'Legacy policy' 'OK' 'No user-agent policy, allow list or block list'
    }
    elseif ($Fix -and $Cmdlet.ShouldProcess($where, 'Clear the legacy EWS user-agent policy and lists')) {
        try {
            # Clearing the policy clears both lists with it. Exchange rejects a list change without a policy, and
            # rejects both lists in one call, so this is the one command that works.
            Set-OrganizationConfig -EwsApplicationAccessPolicy $null -ErrorAction Stop
            $after = Get-OrganizationConfig -RetrieveEwsOperationAccessPolicy
            if ("$($after.EwsApplicationAccessPolicy)".Trim() -or (Get-Values $after.EwsAllowList) -or (Get-Values $after.EwsBlockList)) {
                throw 'still set after the change'
            }
            Add-Finding 'Tenant' 'Legacy policy' 'Fixed' "Cleared. Was $was"
        }
        catch { Add-Finding 'Tenant' 'Legacy policy' 'Problem' "Couldn't clear ($was): $($_.Exception.Message)" }
    }
    else {
        # EnforceAllowList only lets through clients whose user agent is listed, so it blocks most apps outright.
        $sev = if ($legacyPolicy -eq 'EnforceAllowList') { 'Problem' } else { 'Warning' }
        Add-Finding 'Tenant' 'Legacy policy' $sev "$was. These match user agents, not app IDs, and can still block your tool"
    }

    # ---------------------------------------------------------- 3. mailboxes with EWS off on their own
    $cas = Get-CASMailbox -ResultSize Unlimited
    $blocked = @($cas | Where-Object { $_.EwsEnabled -eq $false -or "$($_.EwsApplicationAccessPolicy)".Trim() -or (Get-Values $_.EwsAllowList) -or (Get-Values $_.EwsBlockList) })
    if (-not $blocked) { Add-Finding 'Mailboxes' "$(@($cas).Count) checked" 'OK' 'None block EWS on their own' }
    foreach ($c in $blocked) {
        $why = "EwsEnabled '$($c.EwsEnabled)', policy '$($c.EwsApplicationAccessPolicy)', allow '$((Get-Values $c.EwsAllowList) -join ';')', block '$((Get-Values $c.EwsBlockList) -join ';')'"
        if ($Fix -and $Cmdlet.ShouldProcess($c.PrimarySmtpAddress, 'Re-enable EWS and clear mailbox-level rules')) {
            try {
                Set-CASMailbox -Identity $c.Identity -EwsEnabled $true -ErrorAction Stop
                if ("$($c.EwsApplicationAccessPolicy)".Trim() -or (Get-Values $c.EwsAllowList) -or (Get-Values $c.EwsBlockList)) {
                    Set-CASMailbox -Identity $c.Identity -EwsApplicationAccessPolicy $null -ErrorAction Stop
                }
                Add-Finding 'Mailbox' $c.PrimarySmtpAddress 'Fixed' "Was $why"
            }
            catch { Add-Finding 'Mailbox' $c.PrimarySmtpAddress 'Problem' "Couldn't fix ($why): $($_.Exception.Message)" }
        }
        else { Add-Finding 'Mailbox' $c.PrimarySmtpAddress 'Problem' $why }
    }

    # ---------------------------------------------------------- 4. does the app have EWS permission here?
    if (-not $SkipConsentCheck) {
        try {
            Connect-MgGraph -TenantId $TenantId -Scopes 'Application.Read.All' -NoWelcome
            if ((Get-MgContext).TenantId -ne $TenantId) { throw "Graph signed in to $((Get-MgContext).TenantId)" }
            $sp  = Get-MgServicePrincipal -Filter "appId eq '$AppId'"
            $exo = Get-MgServicePrincipal -Filter "appId eq '00000002-0000-0ff1-ce00-000000000000'"   # Office 365 Exchange Online
            if (-not $sp) { Add-Finding 'Consent' $AppId 'Problem' 'No service principal. The app was never consented in this tenant' }
            else {
                $roles = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -All | Where-Object ResourceId -eq $exo.Id |
                    ForEach-Object { $id = $_.AppRoleId; ($exo.AppRoles | Where-Object Id -eq $id).Value })
                if ($roles -contains 'full_access_as_app') { Add-Finding 'Consent' $sp.DisplayName 'OK' 'Has full_access_as_app' }
                else { Add-Finding 'Consent' $sp.DisplayName 'Warning' "No full_access_as_app (has: '$($roles -join ', ')'). Needed if your tool uses app-only EWS; tools that use delegated EWS don't need it. Re-consent as a Global Admin if yours is app-only" }
            }
        }
        catch { Add-Finding 'Consent' $AppId 'Warning' "Couldn't check: $($_.Exception.Message)" }
    }

    # ---------------------------------------------------------- 5. are the archives actually ready?
    if ($CheckArchives) {
        $active = @(Get-EXOMailbox -ResultSize Unlimited -Properties ArchiveStatus | Where-Object ArchiveStatus -eq 'Active')
        $notReady = @($active | Where-Object { -not (Get-EXOMailboxStatistics -Identity $_.ExternalDirectoryObjectId -Archive -ErrorAction SilentlyContinue) })
        Add-Finding 'Archives' "$($active.Count) active" $(if ($notReady) { 'Warning' } else { 'OK' }) "$($active.Count - $notReady.Count) ready, $($notReady.Count) not ready yet"
        foreach ($n in $notReady) {
            Add-Finding 'Archive' $n.PrimarySmtpAddress 'Warning' 'Active but not readable yet. New archives can take hours. Re-check before copying'
        }
    }
}

# ============================================================ run
if ($Only -ne 'Destination') { Test-Side 'Source' $SourceTenantId $SourceAdmin }
if ($Only -ne 'Source')      { Test-Side 'Destination' $DestinationTenantId $DestinationAdmin }
Disconnect-All

$file = ".\EwsArchiveReadiness-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
$findings | Export-Csv $file -NoTypeInformation -WhatIf:$false

Write-Host ''
Write-Host '==================== SUMMARY ====================' -ForegroundColor Cyan
$findings | Group-Object Side | Sort-Object { $_.Name -ne 'Source' } | ForEach-Object {
    $s = $_.Group
    $line = '{0,-12} {1,-28} problems {2}, fixed {3}, warnings {4}' -f $_.Name, $s[-1].Tenant,
        @($s | Where-Object Status -eq 'Problem').Count, @($s | Where-Object Status -eq 'Fixed').Count, @($s | Where-Object Status -eq 'Warning').Count
    Write-Host $line
}
Write-Host "Report: $file" -ForegroundColor Cyan
$problems = @($findings | Where-Object Status -eq 'Problem').Count
$fixed    = @($findings | Where-Object Status -eq 'Fixed').Count
if ($problems) { Write-Host "$problems problem(s) left. Re-run with -Fix (try -WhatIf first). Archives won't copy until these are cleared." -ForegroundColor Red }
elseif ($fixed) { Write-Host 'Fixes applied. Allow time (Microsoft says up to 24 hours), restart your migration tool, then re-run without -Fix to confirm.' -ForegroundColor Cyan }
else { Write-Host 'No blockers in either tenant. Archives can copy over EWS.' -ForegroundColor Green }
