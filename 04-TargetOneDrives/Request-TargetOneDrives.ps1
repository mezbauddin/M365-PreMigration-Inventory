<#
.SYNOPSIS
    Pre-provisions OneDrive for the people in a tenant-to-tenant migration, so the migration tool has somewhere to
    copy to, then reports each OneDrive's address.

.DESCRIPTION
    TL;DR: A OneDrive doesn't exist until the user signs in or an admin asks for it, and migration tools can't copy
    into one that isn't there. This script reads the people from TargetBuildPlan.csv (or a list you pass), asks
    SharePoint to create their OneDrives in batches of 200, optionally waits until they exist, and writes a report
    with each URL. Users who already have a OneDrive are reported, not re-requested.

.PARAMETER AdminUrl
    The destination's SharePoint admin URL, e.g. https://contoso-admin.sharepoint.com

.PARAMETER PlanFile
    TargetBuildPlan.csv from New-TargetMailboxes. Rows with Include True and Kind User are used. Default
    .\TargetBuildPlan.csv

.PARAMETER User
    Sign-in names (UPNs) to provision instead of, or as well as, the plan file.

.PARAMETER Wait
    Keep checking until every OneDrive exists or -TimeoutMinutes passes.

.PARAMETER TimeoutMinutes
    How long -Wait keeps checking. Default 60.

.PARAMETER AdminUpn
    Optional destination admin sign-in name, e.g. admin@contoso.com. Exchange Online is connected as this account
    (disconnecting any session for another account first) so addresses are resolved to sign-in names in the right
    tenant. Without it, an open Exchange Online session is reused as it is.

.EXAMPLE
    .\Request-TargetOneDrives.ps1 -AdminUrl https://contoso-admin.sharepoint.com

.EXAMPLE
    .\Request-TargetOneDrives.ps1 -AdminUrl https://contoso-admin.sharepoint.com -AdminUpn admin@contoso.com -Wait

.EXAMPLE
    .\Request-TargetOneDrives.ps1 -AdminUrl https://contoso-admin.sharepoint.com -User kim@contoso.com -Wait

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.1.0
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: Microsoft.Online.SharePoint.PowerShell, SharePoint Administrator in the destination. Optional:
              connect ExchangeOnlineManagement first so mail addresses are resolved to sign-in names. Users need a
              licence that includes OneDrive. Supports -WhatIf.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][string]$AdminUrl,
    [string]$PlanFile = '.\TargetBuildPlan.csv',
    [string[]]$User = @(),
    [switch]$Wait,
    [int]$TimeoutMinutes = 60,
    [string]$AdminUpn
)

$ErrorActionPreference = 'Stop'
if ($AdminUrl -notmatch '^https://') { $AdminUrl = "https://$AdminUrl" }

# ------------------------------------------------------------ module and sign-in
if (-not (Get-Command Connect-SPOService -ErrorAction SilentlyContinue)) {
    try { Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -ErrorAction Stop }
    catch {
        if ($PSVersionTable.PSEdition -eq 'Core') {
            # Older versions of the module only load through Windows PowerShell.
            Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -DisableNameChecking -WarningAction SilentlyContinue
        }
        else { throw 'Microsoft.Online.SharePoint.PowerShell is not installed. Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser' }
    }
}
Connect-SPOService -Url $AdminUrl

# ------------------------------------------------------------ who
$people = [System.Collections.Generic.List[string]]::new()
if (Test-Path $PlanFile) {
    Import-Csv $PlanFile | Where-Object { $_.Include -eq 'True' -and $_.Kind -eq 'User' } |
        ForEach-Object { $people.Add($_.TargetAddress.ToLower()) }
}
$User | ForEach-Object { $people.Add($_.ToLower()) }
$people = @($people | Select-Object -Unique)
if (-not $people) { throw "No users. Pass -User, or run from the folder with $PlanFile." }

# With -AdminUpn, Exchange Online is (re)connected as that account, so the look-up below can't run in the source.
if ($AdminUpn) {
    $c = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object State -eq 'Connected' | Select-Object -First 1
    if ($c -and "$($c.UserPrincipalName)" -ne $AdminUpn) {
        Write-Host "Exchange Online is signed in as $($c.UserPrincipalName). Reconnecting as $AdminUpn." -ForegroundColor Yellow
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        $c = $null
    }
    if (-not $c) { Connect-ExchangeOnline -UserPrincipalName $AdminUpn -ShowBanner:$false }
    $c = Get-ConnectionInformation | Where-Object State -eq 'Connected' | Select-Object -First 1
    if ("$($c.UserPrincipalName)" -ne $AdminUpn) { throw "Exchange Online signed in as $($c.UserPrincipalName), not $AdminUpn." }
    Write-Host "Signed in as $($c.UserPrincipalName) to tenant $($c.TenantID)" -ForegroundColor Cyan
}

# A OneDrive belongs to the sign-in name (UPN), which isn't always the mail address in the plan. If Exchange Online
# is connected, look each one up. Otherwise use the addresses as given.
if ((Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) -and
    (Get-ConnectionInformation | Where-Object State -eq 'Connected')) {
    $people = @(foreach ($p in $people) {
        $m = Get-EXOMailbox -Identity $p -Properties UserPrincipalName -ErrorAction SilentlyContinue
        if ($m -and $m.UserPrincipalName.ToLower() -ne $p) {
            Write-Warning "$p signs in as $($m.UserPrincipalName). Using that."
            $m.UserPrincipalName.ToLower()
        }
        else { $p }
    }) | Select-Object -Unique
    $people = @($people)
}
else {
    Write-Warning 'Exchange Online is not connected, so addresses are used as sign-in names. Connect-ExchangeOnline first if they might differ.'
}
Write-Host "$($people.Count) user(s) to check." -ForegroundColor Cyan

function Get-PersonalPath($upn) { '/personal/' + ($upn.ToLower() -replace '[@.]', '_') }

function Get-OneDriveMap {
    # Brand-new OneDrives often have no Owner yet, so match on the URL as well.
    $sites = @(Get-SPOSite -IncludePersonalSite $true -Limit All -Filter "Url -like '-my.sharepoint.com/personal/'")
    $map = @{}
    foreach ($p in $people) {
        $path = Get-PersonalPath $p
        $hit = $sites | Where-Object { ($_.Owner -and $_.Owner.ToLower() -eq $p) -or $_.Url.ToLower().EndsWith($path) } | Select-Object -First 1
        if ($hit) { $map[$p] = $hit.Url }
    }
    return $map
}

# ------------------------------------------------------------ request the missing ones
$existing = Get-OneDriveMap
$missing = @($people | Where-Object { -not $existing.ContainsKey($_) })
Write-Host "$($people.Count - $missing.Count) already have a OneDrive. $($missing.Count) to request." -ForegroundColor Cyan

for ($i = 0; $i -lt $missing.Count; $i += 200) {
    # Request-SPOPersonalSite takes at most 200 users per call.
    $batch = $missing[$i..([Math]::Min($i + 199, $missing.Count - 1))]
    if ($PSCmdlet.ShouldProcess("$($batch.Count) user(s)", 'Request OneDrive')) {
        Request-SPOPersonalSite -UserEmails $batch -NoWait
        Write-Host "Requested $($batch.Count) OneDrive(s)." -ForegroundColor Green
    }
}

# ------------------------------------------------------------ wait and report
$deadline = (Get-Date).AddMinutes($TimeoutMinutes)
do {
    $map = Get-OneDriveMap
    $pending = @($people | Where-Object { -not $map.ContainsKey($_) })
    if (-not $Wait -or -not $pending -or $WhatIfPreference) { break }
    Write-Host "$($pending.Count) still being created. Checking again in 5 minutes ($(Get-Date -Format HH:mm))." -ForegroundColor Yellow
    Start-Sleep -Seconds 300
} while ((Get-Date) -lt $deadline)

$report = foreach ($p in $people) {
    [pscustomobject]@{
        User   = $p
        Status = if ($map.ContainsKey($p)) { if ($existing.ContainsKey($p)) { 'Already existed' } else { 'Created' } } else { 'Pending' }
        Url    = $map[$p]
    }
}
$file = ".\TargetOneDrives-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
$report | Export-Csv $file -NoTypeInformation -WhatIf:$false
$report | Format-Table -AutoSize | Out-Host
Write-Host "Report: $file" -ForegroundColor Cyan
$left = @($report | Where-Object Status -eq 'Pending').Count
if ($left) { Write-Host "$left OneDrive(s) still pending. Creation can take up to 24 hours. Re-run later to check; it won't request them twice." -ForegroundColor Yellow }
else { Write-Host 'Every OneDrive exists. Copy once sign-in names are final, because the URL follows the UPN.' -ForegroundColor Green }
