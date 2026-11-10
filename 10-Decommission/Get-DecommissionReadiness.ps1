<#
.SYNOPSIS
    Read-only check of the SOURCE tenant before you remove licences and let it close: who can still sign in, what's
    still licensed and when it renews, what's on hold, which OneDrives are unlicensed or archived, which domains and
    mailboxes are left, and whether mail is still flowing. PASS, WARN and INFO for each, plus CSVs.

.DESCRIPTION
    TL;DR: Run it at T+7, again before you remove licences (T+30), and again 14 days before the licences end. It
    signs in to the source with Graph and Exchange Online (and SharePoint Online if you give -SharePointAdminUrl) and
    reports, without changing anything:

      Sign-in        Enabled member accounts (your -AdminUpn and -Exclude accounts don't count)
      Licences       Users with licences, and each subscription's status and next renewal or end date
      Domains        Custom domains still in the source (after cutover there should be none)
      Holds          Litigation holds, In-Place and retention holds on mailboxes, organisation-wide holds,
                     inactive mailboxes
      Mailboxes      How many are left, by type, and their total size
      Mail flow      Messages sent or delivered in the last 7 days (message trace). Anything here means someone or
                     something still uses the source
      OneDrive       With -SharePointAdminUrl: OneDrives whose owner has no licence, their size, lock and archive state

    Writes DecommissionReadiness-<date>.csv (Area, Check, Result, Detail, Fix) and detail files
    DecommissionDetail-<kind>-<date>.csv for users, subscriptions, mailboxes and OneDrives.

    Tip: a WARN isn't a stop sign. It's a decision for the business owner. Read each one with them and write the
    answer in the decommission plan.
    Trap: a hold (litigation, retention or eDiscovery) means someone decided that data must be kept. Removing the
    licence or deleting the user doesn't release it, and closing the tenant would. Get the hold owner's decision first.

.PARAMETER TenantId
    Source tenant ID (GUID). The script stops if Graph or Exchange Online is signed in anywhere else.

.PARAMETER AdminUpn
    Optional source admin sign-in name, e.g. admin@fabrikam.onmicrosoft.com. Any session for another account or
    tenant is disconnected first, and the script prints who and where. Not counted as an enabled user.

.PARAMETER SharePointAdminUrl
    Optional source SharePoint admin URL, e.g. https://fabrikam-admin.sharepoint.com. Adds the OneDrive checks.
    Needs Microsoft.Online.SharePoint.PowerShell.

.PARAMETER Exclude
    UPNs that are expected to stay enabled until the end, such as a break-glass Global Administrator.

.PARAMETER MailFlowDays
    How many days of message trace to read. Default 7, most 10.

.PARAMETER RenewalWarningDays
    Subscriptions renewing or ending within this many days are a WARN. Default 30.

.PARAMETER SkipMailFlow
    Don't run the message trace.

.EXAMPLE
    .\Get-DecommissionReadiness.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com

.EXAMPLE
    .\Get-DecommissionReadiness.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com `
        -SharePointAdminUrl https://fabrikam-admin.sharepoint.com -Exclude breakglass@fabrikam.onmicrosoft.com

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.0.0
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: Microsoft.Graph.Authentication, ExchangeOnlineManagement 3.x. Optional:
              Microsoft.Online.SharePoint.PowerShell. Global Reader is enough for everything except message trace
              (Exchange Administrator or a role with Message Tracking) and SharePoint (SharePoint Administrator).
    Changes nothing in any tenant.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F-]{36}$')][string]$TenantId,
    [string]$AdminUpn,
    [ValidatePattern('^(https://)?[a-z0-9-]+-admin\.sharepoint\.com/?$')][string]$SharePointAdminUrl,
    [string[]]$Exclude = @(),
    [ValidateRange(1, 10)][int]$MailFlowDays = 7,
    [ValidateRange(1, 365)][int]$RenewalWarningDays = 30,
    [switch]$SkipMailFlow
)

$ErrorActionPreference = 'Stop'
foreach ($m in 'Microsoft.Graph.Authentication', 'ExchangeOnlineManagement') {
    if (-not (Get-Module -ListAvailable $m)) { throw "Install-Module $m -Scope CurrentUser first." }
}
$stamp = Get-Date -Format 'yyyyMMdd-HHmm'

$results = [System.Collections.Generic.List[object]]::new()
function Add-Result($Area, $Check, $Result, $Detail, $Fix = '') {
    $results.Add([pscustomobject]@{ Area = $Area; Check = $Check; Result = $Result; Detail = $Detail; Fix = $Fix })
    $c = @{ PASS = 'Green'; WARN = 'Yellow'; INFO = 'Gray'; SKIP = 'Gray' }[$Result]
    Write-Host ("  {0,-4} {1}: {2}" -f $Result, $Check, $Detail) -ForegroundColor $c
}
function Save-Detail([string]$Kind, $Rows) {
    if (@($Rows).Count) { @($Rows) | Export-Csv ".\DecommissionDetail-$Kind-$stamp.csv" -NoTypeInformation }
}
function Get-GraphAll([string]$Uri) {
    $out = [System.Collections.Generic.List[object]]::new()
    while ($Uri) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $Uri
        foreach ($v in @($page.value)) { $out.Add($v) }
        $Uri = $page.'@odata.nextLink'
    }
    $out
}
function Get-First([string[]]$Items, [int]$Count = 8) {
    $list = @($Items | Where-Object { $_ })
    $text = ($list | Select-Object -First $Count) -join ', '
    if ($list.Count -gt $Count) { $text += " and $($list.Count - $Count) more" }
    $text
}

# ------------------------------------------------------------ Graph: sign in to the right tenant as the right account
$ctx = Get-MgContext
if ($ctx -and ($ctx.TenantId -ne $TenantId -or ($AdminUpn -and $ctx.Account -ne $AdminUpn))) {
    Write-Host "Graph is signed in as $($ctx.Account) to $($ctx.TenantId). Disconnecting." -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
}
Connect-MgGraph -TenantId $TenantId -Scopes 'User.Read.All', 'Directory.Read.All', 'Organization.Read.All', 'LicenseAssignment.Read.All' -NoWelcome
$ctx = Get-MgContext
if ($ctx.TenantId -ne $TenantId) { throw "Graph signed in to $($ctx.TenantId), not $TenantId. Stopping." }
if ($AdminUpn -and $ctx.Account -ne $AdminUpn) { throw "Graph signed in as $($ctx.Account), not $AdminUpn. Run Disconnect-MgGraph and pick $AdminUpn." }
Write-Host "Signed in as $($ctx.Account) to tenant $($ctx.TenantId)" -ForegroundColor Cyan
$skip = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
@($Exclude) + $AdminUpn + $ctx.Account | Where-Object { $_ } | ForEach-Object { [void]$skip.Add($_.Trim()) }

# ------------------------------------------------------------ sign-in and licences
Write-Host 'Users and licences' -ForegroundColor Cyan
$skus = @{}
foreach ($s in @(Get-GraphAll 'v1.0/subscribedSkus')) { $skus["$($s.skuId)"] = $s }
$users = @(Get-GraphAll 'v1.0/users?$select=id,displayName,userPrincipalName,accountEnabled,userType,assignedLicenses&$top=999')
$members = @($users | Where-Object { $_.userType -ne 'Guest' })
$enabled = @($members | Where-Object { $_.accountEnabled -and -not $skip.Contains($_.userPrincipalName) })
$licensedUsers = @($users | Where-Object { @($_.assignedLicenses).Count })
Save-Detail 'Users' ($members | ForEach-Object {
        [pscustomobject]@{
            DisplayName       = $_.displayName
            UserPrincipalName = $_.userPrincipalName
            Enabled           = $_.accountEnabled
            Excluded          = $skip.Contains($_.userPrincipalName)
            Licences          = (@($_.assignedLicenses) | ForEach-Object { $skus["$($_.skuId)"].skuPartNumber }) -join ';'
        } })
if ($enabled) { Add-Result 'Sign-in' 'Enabled accounts' 'WARN' "$($enabled.Count) can still sign in: $(Get-First $enabled.userPrincipalName)" 'Block them with 07 Block-SourceSignIn.ps1. Admin-role accounts in the admin centre, last. Add a break-glass account to -Exclude' }
else { Add-Result 'Sign-in' 'Enabled accounts' 'PASS' "None, apart from $($skip.Count) excluded or your own" }

if ($licensedUsers) {
    $bySku = $licensedUsers | ForEach-Object { foreach ($l in @($_.assignedLicenses)) { $skus["$($l.skuId)"].skuPartNumber } } | Group-Object | ForEach-Object { "$($_.Name) x$($_.Count)" }
    Add-Result 'Licences' 'Licensed users' 'WARN' "$($licensedUsers.Count) user(s) still licensed ($($bySku -join ', '))" 'Remove at T+30, once exports are done and the OneDrive decision is made. Unlicensed OneDrives then start their timer'
}
else { Add-Result 'Licences' 'Licensed users' 'PASS' 'No user holds a licence' }

# Subscriptions with renewal or end dates. directory/subscriptions has nextLifecycleDateTime; subscribedSkus doesn't.
try {
    $subs = @(Get-GraphAll 'v1.0/directory/subscriptions')
    $subRows = foreach ($s in $subs) {
        $next = if ($s.nextLifecycleDateTime) { [datetime]$s.nextLifecycleDateTime } else { $null }
        $sku = $skus["$($s.skuId)"]
        [pscustomobject]@{
            SkuPartNumber = $s.skuPartNumber
            # Not called Status, so tools/Get-MigrationReport doesn't read this detail file as a report.
            SubscriptionStatus = $s.status
            Trial         = $s.isTrial
            TotalLicences = $s.totalLicenses
            Consumed      = $(if ($sku) { $sku.consumedUnits } else { '' })
            Created       = $(if ($s.createdDateTime) { ([datetime]$s.createdDateTime).ToString('yyyy-MM-dd') } else { '' })
            NextLifecycle = $(if ($next) { $next.ToString('yyyy-MM-dd') } else { '' })
            DaysLeft      = $(if ($next) { [int]($next - (Get-Date)).TotalDays } else { '' })
        }
    }
    $subRows = @($subRows)
    Save-Detail 'Subscriptions' $subRows
    foreach ($s in $subRows | Where-Object { $_.SubscriptionStatus -ne 'Deleted' }) {
        $label = "$($s.SkuPartNumber) ($($s.TotalLicences) seats, $($s.Consumed) used)"
        if ($s.SubscriptionStatus -in 'Warning', 'Suspended', 'LockedOut') {
            Add-Result 'Licences' $label 'WARN' "Status $($s.SubscriptionStatus), next change $($s.NextLifecycle)" 'Already lapsing: data is on Microsoft''s timetable now. Check exports are finished'
        }
        elseif ("$($s.DaysLeft)" -ne '' -and [int]$s.DaysLeft -le $RenewalWarningDays) {
            Add-Result 'Licences' $label 'WARN' "Renews or ends on $($s.NextLifecycle), in $($s.DaysLeft) day(s)" 'Decide now: turn off recurring billing or cancel, or it renews for another term'
        }
        else { Add-Result 'Licences' $label 'INFO' "Status $($s.SubscriptionStatus), renews or ends on $(if ($s.NextLifecycle) { $s.NextLifecycle } else { 'unknown' })" }
    }
}
catch {
    Add-Result 'Licences' 'Subscription dates' 'INFO' "Couldn't read directory/subscriptions ($($_.Exception.Message)). Check renewal dates in the admin centre: Billing, Your products"
}

# ------------------------------------------------------------ domains
$domains = @((Invoke-MgGraphRequest GET 'v1.0/domains').value)
$custom = @($domains | Where-Object { $_.id -notmatch '\.onmicrosoft\.com$' })
if ($custom) { Add-Result 'Domains' 'Custom domains' 'WARN' "Still in the source: $(Get-First $custom.id)" 'Anything moved should already be gone (07 Remove-SourceDomain). Remove the rest before the tenant closes, or it stays tied up' }
else { Add-Result 'Domains' 'Custom domains' 'PASS' 'Only onmicrosoft.com domains left' }

# ------------------------------------------------------------ Exchange Online
$c = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object State -eq 'Connected' | Select-Object -First 1
if ($c -and ("$($c.TenantID)" -ne $TenantId -or ($AdminUpn -and "$($c.UserPrincipalName)" -ne $AdminUpn))) {
    Write-Host "Exchange Online is signed in as $($c.UserPrincipalName) to $($c.TenantID). Disconnecting." -ForegroundColor Yellow
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    $c = $null
}
if (-not $c) {
    $p = @{ ShowBanner = $false }
    if ($AdminUpn) { $p.UserPrincipalName = $AdminUpn }
    Connect-ExchangeOnline @p
    $c = Get-ConnectionInformation | Where-Object State -eq 'Connected' | Select-Object -First 1
}
if ("$($c.TenantID)" -ne $TenantId) { throw "Exchange Online is signed in to tenant $($c.TenantID), not $TenantId. Stopping." }
Write-Host "Exchange Online: $($c.UserPrincipalName)" -ForegroundColor Cyan

Write-Host 'Mailboxes and holds' -ForegroundColor Cyan
$mbx = @(Get-EXOMailbox -ResultSize Unlimited -Properties LitigationHoldEnabled, InPlaceHolds, ArchiveStatus | Where-Object RecipientTypeDetails -ne 'DiscoveryMailbox')
$mbxRows = foreach ($m in $mbx) {
    $st = Get-EXOMailboxStatistics -Identity $m.ExternalDirectoryObjectId -ErrorAction SilentlyContinue
    $bytes = if ($st -and "$($st.TotalItemSize)" -match '\(([\d,]+) bytes\)') { [long]($Matches[1] -replace ',', '') } else { [long]0 }
    [pscustomobject]@{
        PrimarySmtpAddress   = "$($m.PrimarySmtpAddress)"
        RecipientTypeDetails = $m.RecipientTypeDetails
        SizeGB               = [math]::Round($bytes / 1e9, 2)
        ArchiveStatus        = $m.ArchiveStatus
        LitigationHold       = $m.LitigationHoldEnabled
        InPlaceHolds         = (@($m.InPlaceHolds) -join ';')
        LastLogon            = $(if ($st) { $st.LastLogonTime } else { '' })
    }
}
$mbxRows = @($mbxRows)
Save-Detail 'Mailboxes' $mbxRows
if ($mbxRows) {
    $types = ($mbxRows | Group-Object RecipientTypeDetails | ForEach-Object { "$($_.Name) $($_.Count)" }) -join ', '
    Add-Result 'Mailboxes' 'Mailboxes left' 'INFO' ("{0} mailbox(es), {1} GB primary ({2})" -f $mbxRows.Count, [math]::Round(($mbxRows | Measure-Object SizeGB -Sum).Sum, 2), $types) 'Each one is deleted with its licence and user. Export anything the business must keep first'
}
else { Add-Result 'Mailboxes' 'Mailboxes left' 'PASS' 'None' }

$held = @($mbxRows | Where-Object { $_.LitigationHold -eq $true -or $_.InPlaceHolds })
if ($held) { Add-Result 'Holds' 'Mailboxes on hold' 'WARN' "$($held.Count): $(Get-First $held.PrimarySmtpAddress)" 'Data someone decided to keep. Get the hold owner to release it, or export it (eDiscovery) before the tenant closes' }
else { Add-Result 'Holds' 'Mailboxes on hold' 'PASS' 'No litigation, In-Place or retention hold on any mailbox' }

$orgHolds = @((Get-OrganizationConfig).InPlaceHolds | Where-Object { $_ })
if ($orgHolds) { Add-Result 'Holds' 'Organisation-wide holds' 'WARN' "$($orgHolds.Count) retention or hold policy entries apply to every mailbox" 'Check Purview, Data lifecycle management and eDiscovery cases. Decide with the policy owner' }
else { Add-Result 'Holds' 'Organisation-wide holds' 'PASS' 'None' }

$inactive = @()
try { $inactive = @(Get-Mailbox -InactiveMailboxOnly -ResultSize Unlimited -ErrorAction Stop) } catch { Write-Warning "Couldn't read inactive mailboxes: $($_.Exception.Message)" }
if ($inactive) { Add-Result 'Holds' 'Inactive mailboxes' 'WARN' "$($inactive.Count) mailbox(es) of deleted users kept by a hold: $(Get-First ($inactive | ForEach-Object { "$($_.PrimarySmtpAddress)" }))" 'They go when the tenant goes. Export them with eDiscovery if they must be kept' }
else { Add-Result 'Holds' 'Inactive mailboxes' 'PASS' 'None' }

# ------------------------------------------------------------ mail flow
if ($SkipMailFlow) { Add-Result 'Mail flow' "Last $MailFlowDays days" 'SKIP' 'Skipped (-SkipMailFlow)' }
else {
    Write-Host "Message trace, last $MailFlowDays days" -ForegroundColor Cyan
    $accepted = @((Get-AcceptedDomain).DomainName | ForEach-Object { "$_".ToLower() })
    $start = (Get-Date).AddDays(-$MailFlowDays); $end = Get-Date
    $trace = $null
    try {
        # Get-MessageTraceV2 replaced Get-MessageTrace. It returns up to 5000 rows and spans at most 10 days.
        if (Get-Command Get-MessageTraceV2 -ErrorAction SilentlyContinue) { $trace = @(Get-MessageTraceV2 -StartDate $start -EndDate $end -ResultSize 5000 -ErrorAction Stop) }
        else { $trace = @(Get-MessageTrace -StartDate $start -EndDate $end -PageSize 5000 -ErrorAction Stop) }
    }
    catch { Add-Result 'Mail flow' "Last $MailFlowDays days" 'INFO' "Message trace failed: $($_.Exception.Message)" 'Run it from the Exchange admin centre: Mail flow, Message trace' }
    if ($null -ne $trace) {
        $isOurs = { param($a) $d = "$a".Split('@')[-1].ToLower(); $accepted -contains $d }
        $inbound = @($trace | Where-Object { "$($_.Status)" -eq 'Delivered' -and (& $isOurs $_.RecipientAddress) })
        $outbound = @($trace | Where-Object { (& $isOurs $_.SenderAddress) -and -not (& $isOurs $_.RecipientAddress) })
        $capped = if ($trace.Count -ge 5000) { ' (5000 rows read: there may be more)' } else { '' }
        if ($inbound) { Add-Result 'Mail flow' 'Delivered into the source' 'WARN' "$($inbound.Count) message(s)$capped. Top recipients: $(Get-First ($inbound | Group-Object RecipientAddress | Sort-Object Count -Descending | ForEach-Object { "$($_.Name) ($($_.Count))" }) 5)" 'Something still sends here: a contact form, scanner, app or a sender with an old onmicrosoft.com address. Find and repoint it' }
        else { Add-Result 'Mail flow' 'Delivered into the source' 'PASS' "Nothing delivered in $MailFlowDays days" }
        if ($outbound) { Add-Result 'Mail flow' 'Sent from the source' 'WARN' "$($outbound.Count) message(s)$capped. Top senders: $(Get-First ($outbound | Group-Object SenderAddress | Sort-Object Count -Descending | ForEach-Object { "$($_.Name) ($($_.Count))" }) 5)" 'Someone, or an app or device, is still sending from the old tenant. Block it' }
        else { Add-Result 'Mail flow' 'Sent from the source' 'PASS' "Nothing sent in $MailFlowDays days" }
    }
}

# ------------------------------------------------------------ OneDrive (optional)
if (-not $SharePointAdminUrl) {
    Add-Result 'OneDrive' 'Unlicensed OneDrives' 'SKIP' 'No -SharePointAdminUrl. Check SharePoint admin centre, Reports, OneDrive accounts'
}
else {
    $SharePointAdminUrl = 'https://' + ($SharePointAdminUrl.Trim().ToLower() -replace '^https://', '' -replace '/$', '')
    try {
        if (-not (Get-Command Connect-SPOService -ErrorAction SilentlyContinue)) {
            if (-not (Get-Module -ListAvailable Microsoft.Online.SharePoint.PowerShell)) { throw 'Microsoft.Online.SharePoint.PowerShell is not installed' }
            if ($PSVersionTable.PSEdition -eq 'Core') { Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -DisableNameChecking -WarningAction SilentlyContinue }
            else { Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -WarningAction SilentlyContinue }
        }
        Write-Host "SharePoint Online ($SharePointAdminUrl)" -ForegroundColor Cyan
        Connect-SPOService -Url $SharePointAdminUrl
        $licensedSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $licensedUsers | ForEach-Object { [void]$licensedSet.Add($_.userPrincipalName) }
        $drives = @(Get-SPOSite -IncludePersonalSite $true -Limit All -Filter "Url -like '-my.sharepoint.com/personal/'")
        $driveRows = foreach ($d in $drives) {
            # Archive state is exposed as ArchiveStatus on current module versions; fall back to blank on older ones.
            $archive = foreach ($n in 'ArchiveStatus', 'ArchiveState') { if ($d.PSObject.Properties[$n] -and "$($d.$n)") { "$($d.$n)"; break } }
            [pscustomobject]@{
                Url        = $d.Url
                Owner      = $d.Owner
                Licensed   = $licensedSet.Contains("$($d.Owner)")
                StorageGB  = [math]::Round($d.StorageUsageCurrent / 1024, 2)
                LockState  = $d.LockState
                Archive    = "$archive"
                LastChange = $d.LastContentModifiedDate
            }
        }
        $driveRows = @($driveRows)
        Save-Detail 'OneDrives' $driveRows
        $unlicensed = @($driveRows | Where-Object { -not $_.Licensed -and $_.StorageGB -gt 0 })
        $archived = @($driveRows | Where-Object { $_.Archive -and $_.Archive -notmatch '^(NotArchived|None|Active)$' })
        if ($unlicensed) {
            Add-Result 'OneDrive' 'Unlicensed OneDrives' 'WARN' ("{0} with data, {1} GB, owner unlicensed: {2}" -f $unlicensed.Count, [math]::Round(($unlicensed | Measure-Object StorageGB -Sum).Sum, 2), (Get-First $unlicensed.Owner 5)) 'Decide per person before licence end: copied already, relicense and copy, eDiscovery export, or pay-as-you-go archive. Read-only at day 60, archived at day 93 by default'
        }
        else { Add-Result 'OneDrive' 'Unlicensed OneDrives' 'PASS' 'Every OneDrive with data has a licensed owner, or there are none' }
        if ($archived) { Add-Result 'OneDrive' 'Archived OneDrives' 'WARN' "$($archived.Count) archived or archiving: $(Get-First $archived.Owner 5)" 'No copy tool can read these until reactivated: relicense the owner (free, 24 to 48 hours) or reactivate (pay-as-you-go, per GB)' }
        elseif ($driveRows) { Add-Result 'OneDrive' 'Archived OneDrives' 'PASS' 'None archived' }
        $sites = @(Get-SPOSite -Limit All)
        if ($sites) { Add-Result 'SharePoint' 'Sites left' 'INFO' ("{0} site(s), {1} GB" -f $sites.Count, [math]::Round(($sites | Measure-Object StorageUsageCurrent -Sum).Sum / 1024, 2)) 'Deleted with the tenant. Check everything needed was copied (06, copy options)' }
        Disconnect-SPOService -ErrorAction SilentlyContinue
    }
    catch { Add-Result 'OneDrive' 'Unlicensed OneDrives' 'SKIP' "SharePoint check failed: $($_.Exception.Message)" 'Check SharePoint admin centre, Reports, OneDrive accounts' }
}

# ------------------------------------------------------------ summary
$file = ".\DecommissionReadiness-$stamp.csv"
$results | Export-Csv $file -NoTypeInformation
Write-Host ''
$results | Format-Table Area, Check, Result, Detail -AutoSize -Wrap | Out-String -Width 250 | Write-Host
$warns = @($results | Where-Object Result -eq 'WARN').Count
if ($warns) { Write-Host "$warns WARN. Each one is a decision for the business owner before licences are removed. Report: $file" -ForegroundColor Yellow }
else { Write-Host "No warnings. Ready for the next decommission step. Report: $file" -ForegroundColor Green }
Write-Host 'Detail files: DecommissionDetail-*.csv. They list people and addresses: keep them out of source control.' -ForegroundColor Cyan
