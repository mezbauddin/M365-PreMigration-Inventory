<#
.SYNOPSIS
    Read-only check of the destination tenant against the source inventory and the mapping file. Writes a pass, warn
    and fail report so you can sign off a migration, or see exactly what's left.

.DESCRIPTION
    TL;DR: Run against the destination after the copy, and again after cutover. For every row in your mapping file it
    checks the recipient exists with the right type, is licensed where TargetBuildPlan.csv says, has an archive where
    the source had one, and has roughly as many items as the source did. With -Domain it checks old addresses and
    sign-in names after cutover. With -SourceSnapshot it checks X500 addresses. With -SharePointAdminUrl it checks
    every person has a OneDrive. With -MailConfigFolder it compares distribution list member counts. Nothing in the
    tenant is changed.

    Checks, per mapped recipient:
      Exists        The destination recipient resolves
      Type          Recipient type matches the plan (User or Shared) or the source
      Licence       Plan rows with a licence have that SKU assigned
      Archive       An active archive where the source had one
      Address       With -Domain: plan rows have the old address as primary; everyone else has it as an alias
      SignIn        With -Domain: plan rows with a mailbox sign in with their old address
      X500          With -SourceSnapshot: the old LegacyExchangeDN is present as an X500 address
      Items         Mailbox item count within -Tolerance percent of 04-Mailboxes.csv
      ArchiveItems  Archive item count, same rule
      OneDrive      With -SharePointAdminUrl: people (Kind User) have a OneDrive
      Password      People (Kind User) must change their temporary password at next sign-in
    Plus:
      Mapping       Plan rows that aren't in the mapping file
      ListMembers   With -MailConfigFolder: member count of each distribution list against the 05 export
      CA-Location   Conditional Access policies with a location condition. Your own test sign-in from abroad (or
                    through a VPN) can be blocked with error 53003, which looks like a broken migration
      MFA           No Conditional Access policy requires MFA for all users and security defaults are off. A
                    finding to report, not something this script changes

.PARAMETER TenantId
    Optional destination tenant ID (GUID). The script stops if Exchange Online or Graph is signed in to any other
    tenant.

.PARAMETER AdminUpn
    Optional admin sign-in name, e.g. admin@contoso.com. If Exchange Online or Graph is already signed in as someone
    else, or to another tenant, the script disconnects and signs in again as this account.

.PARAMETER InventoryFolder
    The Inventory-<tenant>-<date> folder from Invoke-SourceTenantInventory (part 1). 04-Mailboxes.csv is required.
    Re-run the inventory close to cutover so the item counts are recent.

.PARAMETER MappingFile
    CSV with Source and Destination columns, the same file every other part uses.

.PARAMETER PlanFile
    TargetBuildPlan.csv from New-TargetMailboxes. Default .\TargetBuildPlan.csv. Without it, Type falls back to the
    source type and Licence, SignIn and OneDrive aren't checked.

.PARAMETER Domain
    The domains that moved at cutover, e.g. fabrikam.com. Turns on the Address and SignIn checks. Leave it out before
    cutover.

.PARAMETER SourceSnapshot
    SourceRecipients-<date>.csv from Remove-SourceDomain (part 7). Turns on the X500 check.

.PARAMETER Tolerance
    How far below the source item count, in percent, still passes. Default 5.

.PARAMETER MailConfigFolder
    The MailConfig-<domain>-<date> folder from Copy-MailConfig -Export (part 5). Turns on the ListMembers check.

.PARAMETER SharePointAdminUrl
    Destination SharePoint admin URL, e.g. https://contoso-admin.sharepoint.com. Turns on the OneDrive check.

.PARAMETER CutoverDate
    Optional. When cutover happened. A moved user whose password was changed on or after this date passes the
    Password check even though the change-at-next-sign-in flag is off (they've already signed in and changed it).

.PARAMETER SkipTenantChecks
    Leave out the tenant-wide Conditional Access and MFA checks.

.EXAMPLE
    .\Test-TargetTenant.ps1 -TenantId <dest-guid> -InventoryFolder .\Inventory-fabrikam-20261110-0900 -MappingFile .\UserMapping.csv

.EXAMPLE
    .\Test-TargetTenant.ps1 -TenantId <dest-guid> -InventoryFolder .\Inventory-fabrikam-20261110-0900 -MappingFile .\UserMapping.csv `
        -Domain fabrikam.com -SourceSnapshot .\SourceRecipients-20261110-1800.csv -MailConfigFolder .\MailConfig-fabrikam.com-20261109-1100 `
        -SharePointAdminUrl https://contoso-admin.sharepoint.com

.EXAMPLE
    .\Test-TargetTenant.ps1 -TenantId <dest-guid> -AdminUpn admin@contoso.com -InventoryFolder .\Inventory-fabrikam-20261110-0900 `
        -MappingFile .\UserMapping.csv -Domain fabrikam.com -CutoverDate 2026-11-10

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.1.0
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: ExchangeOnlineManagement 3.x, Microsoft.Graph.Authentication, Microsoft.Graph.Users,
              Microsoft.Graph.Identity.DirectoryManagement. Global Reader is enough (Graph scopes User.Read.All,
              Organization.Read.All, Policy.Read.All). For -SharePointAdminUrl:
              Microsoft.Online.SharePoint.PowerShell and SharePoint Administrator.
    Item counts are a guide, not proof. Copy tools can skip some item types, and new mail arrives after cutover.
    Read every Warn and Fail, and decide.
#>
[CmdletBinding()]
param(
    [ValidatePattern('^[0-9a-fA-F-]{36}$')][string]$TenantId,
    [string]$AdminUpn,
    [Parameter(Mandatory)][string]$InventoryFolder,
    [Parameter(Mandatory)][string]$MappingFile,
    [string]$PlanFile = '.\TargetBuildPlan.csv',
    [string[]]$Domain = @(),
    [string]$SourceSnapshot,
    [ValidateRange(0, 100)][double]$Tolerance = 5,
    [string]$MailConfigFolder,
    [string]$SharePointAdminUrl,
    [datetime]$CutoverDate,
    [switch]$SkipTenantChecks
)

$ErrorActionPreference = 'Stop'
$Domain = @($Domain | ForEach-Object { $_.Trim().ToLower() })

# ------------------------------------------------------------ inputs first, so a typo fails before any sign-in
$mbxFile = Join-Path $InventoryFolder '04-Mailboxes.csv'
if (-not (Test-Path $mbxFile)) { throw "04-Mailboxes.csv not found in $InventoryFolder. Point -InventoryFolder at the folder part 1 wrote." }
if ($SourceSnapshot -and -not (Test-Path $SourceSnapshot)) { throw "$SourceSnapshot not found." }
if ($MailConfigFolder) {
    foreach ($f in 'DistributionGroups.csv', 'DistributionGroupMembers.csv') {
        if (-not (Test-Path (Join-Path $MailConfigFolder $f))) { throw "$f missing from $MailConfigFolder. Point -MailConfigFolder at the folder Copy-MailConfig -Export wrote." }
    }
}

$inv = @{}
foreach ($m in Import-Csv $mbxFile) {
    foreach ($a in "$($m.Aliases)".Split(';')) { if ($a -match '^smtp:(.+)$') { $inv[$Matches[1].ToLower()] = $m } }
    $inv["$($m.PrimarySmtpAddress)".ToLower()] = $m
}
$plan = @{}
if (Test-Path $PlanFile) {
    Import-Csv $PlanFile | Where-Object { $_.Include -eq 'True' } | ForEach-Object { $plan["$($_.SourceAddress)".ToLower()] = $_ }
}
else { Write-Warning "$PlanFile not found. Licence, SignIn and OneDrive checks need it, so they're skipped." }
$snapIndex = @{}
if ($SourceSnapshot) {
    foreach ($s in Import-Csv $SourceSnapshot) {
        foreach ($a in "$($s.EmailAddresses)".Split(';')) { if ($a -match '^smtp:(.+)$') { $snapIndex[$Matches[1].ToLower()] = $s } }
        $snapIndex["$($s.PrimarySmtpAddress)".ToLower()] = $s
    }
}
$mapping = [System.Collections.Generic.List[object]]::new()
$seen = [System.Collections.Generic.HashSet[string]]::new()
foreach ($r in Import-Csv $MappingFile) {
    if (-not $r.Source -or -not $r.Destination) { throw "$MappingFile needs Source and Destination columns." }
    if ($seen.Add($r.Source.Trim().ToLower())) { $mapping.Add([pscustomobject]@{ Source = $r.Source.Trim().ToLower(); Destination = $r.Destination.Trim() }) }
}

# ------------------------------------------------------------ connect and guard
# With -AdminUpn, an open session for another account or tenant is closed first, so Windows can't quietly reuse it.
$conn = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object State -eq 'Connected' | Select-Object -First 1
if ($conn -and $AdminUpn -and (("$($conn.UserPrincipalName)" -ne $AdminUpn) -or ($TenantId -and $conn.TenantID -ne $TenantId))) {
    Write-Host "Exchange Online is signed in as $($conn.UserPrincipalName). Reconnecting as $AdminUpn." -ForegroundColor Yellow
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    $conn = $null
}
if (-not $conn) {
    $exoArgs = @{ ShowBanner = $false }
    if ($AdminUpn) { $exoArgs.UserPrincipalName = $AdminUpn }
    Connect-ExchangeOnline @exoArgs
}
$conn = Get-ConnectionInformation | Where-Object State -eq 'Connected' | Select-Object -First 1
$exoTenant = $conn.TenantID
if ($TenantId -and $exoTenant -ne $TenantId) { throw "Exchange Online is signed in to $exoTenant, not $TenantId. Disconnect-ExchangeOnline and sign in to the destination, or pass -AdminUpn." }
if ($AdminUpn -and "$($conn.UserPrincipalName)" -ne $AdminUpn) { throw "Exchange Online signed in as $($conn.UserPrincipalName), not $AdminUpn." }
$mg = Get-MgContext
if ($mg -and (($mg.TenantId -ne $exoTenant) -or ($AdminUpn -and $mg.Account -ne $AdminUpn))) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
Connect-MgGraph -TenantId $exoTenant -NoWelcome -Scopes 'User.Read.All', 'Organization.Read.All', 'Policy.Read.All'
if ((Get-MgContext).TenantId -ne $exoTenant) { throw "Graph signed in to $((Get-MgContext).TenantId), not $exoTenant." }
if ($AdminUpn -and (Get-MgContext).Account -ne $AdminUpn) { throw "Graph signed in as $((Get-MgContext).Account), not $AdminUpn. Run Disconnect-MgGraph and pick the right account." }
Write-Host "Signed in as $($conn.UserPrincipalName) to tenant $exoTenant" -ForegroundColor Cyan
$tenantName = (Get-AcceptedDomain | Where-Object Default).DomainName
Write-Host "Checking $tenantName against $($mapping.Count) mapped recipients. Read-only." -ForegroundColor Cyan

$skus = @{}
foreach ($s in Get-MgSubscribedSku -All) { $skus["$($s.SkuId)"] = $s.SkuPartNumber }

$oneDrives = $null
if ($SharePointAdminUrl) {
    if ($SharePointAdminUrl -notmatch '^https://') { $SharePointAdminUrl = "https://$SharePointAdminUrl" }
    try {
        if (-not (Get-Command Connect-SPOService -ErrorAction SilentlyContinue)) {
            try { Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -ErrorAction Stop }
            catch {
                if ($PSVersionTable.PSEdition -eq 'Core') { Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -DisableNameChecking -WarningAction SilentlyContinue }
                else { throw }
            }
        }
        Connect-SPOService -Url $SharePointAdminUrl
        $oneDrives = @(Get-SPOSite -IncludePersonalSite $true -Limit All -Filter "Url -like '-my.sharepoint.com/personal/'")
    }
    catch { Write-Warning "Couldn't read OneDrives ($($_.Exception.Message)). The OneDrive check is skipped." }
}

# ------------------------------------------------------------ helpers
$results = [System.Collections.Generic.List[object]]::new()
function Add-Result($Object, $Check, $Result, $Expected, $Actual, $Detail) {
    $results.Add([pscustomobject]@{ Object = $Object; Check = $Check; Result = $Result; Expected = $Expected; Actual = $Actual; Detail = $Detail })
    if ($Result -ne 'Pass') {
        $c = @{ Warn = 'Yellow'; Fail = 'Red' }[$Result]
        Write-Host ("  [{0}] {1} {2}: expected {3}, found {4}. {5}" -f $Result, $Object, $Check, $Expected, $Actual, $Detail) -ForegroundColor $c
    }
}
function Get-DomainOf($a) { "$a".Split('@')[-1].ToLower() }
function Test-Count($Object, $Check, $Source, $Dest) {
    $src = [long]("0$Source"); $dst = [long]("0$Dest")
    if ($src -le 0) { Add-Result $Object $Check 'Pass' $src $dst 'Source was empty'; return }
    $short = 100 * ($src - $dst) / $src
    if ($short -gt $Tolerance) { Add-Result $Object $Check 'Fail' $src $dst ("{0:N1}% fewer than the source. Copy not finished, or items skipped. Check the copy tool's report" -f $short) }
    elseif (-$short -gt $Tolerance) { Add-Result $Object $Check 'Warn' $src $dst ("{0:N1}% more than the source. New mail since the inventory, or duplicates from a re-copy" -f -$short) }
    else { Add-Result $Object $Check 'Pass' $src $dst ("Within {0}%" -f $Tolerance) }
}
function Get-PersonalPath($upn) { '/personal/' + ($upn.ToLower() -replace '[@.]', '_') }

# ------------------------------------------------------------ per mapped recipient
$i = 0
foreach ($w in $mapping) {
    $i++; Write-Progress -Activity 'Validating' -Status $w.Source -PercentComplete (100 * $i / $mapping.Count)
    $src = $w.Source
    $p = $plan[$src]; $m = $inv[$src]; $snap = $snapIndex[$src]

    $dest = Get-EXORecipient -Identity $w.Destination -Properties EmailAddresses, ExternalDirectoryObjectId, RecipientTypeDetails -ErrorAction SilentlyContinue
    if (-not $dest) { Add-Result $src 'Exists' 'Fail' $w.Destination 'nothing' 'Destination recipient not found'; continue }
    Add-Result $src 'Exists' 'Pass' $w.Destination "$($dest.PrimarySmtpAddress)" ''
    $type = "$($dest.RecipientTypeDetails)"
    $id = if ($dest.ExternalDirectoryObjectId) { $dest.ExternalDirectoryObjectId } else { "$($dest.Identity)" }
    $have = @($dest.EmailAddresses | ForEach-Object { "$_".ToLower() })

    # ---------------------------------------------------- type
    if ($p) {
        $want = if ($p.Kind -eq 'User') { 'UserMailbox' } else { 'SharedMailbox' }
        if ($type -eq $want) { Add-Result $src 'Type' 'Pass' $want $type '' }
        elseif ($want -eq 'UserMailbox' -and $type -eq 'SharedMailbox') { Add-Result $src 'Type' 'Fail' $want $type 'Still staged as shared. Run part 2 -Apply without -StageOnly' }
        else { Add-Result $src 'Type' 'Fail' $want $type '' }
    }
    else {
        $want = if ($m) { $m.RecipientTypeDetails } elseif ($snap) { $snap.RecipientTypeDetails } else { $null }
        if (-not $want) { Add-Result $src 'Type' 'Warn' 'a source record' $type 'Not in the inventory or snapshot, so the type is not checked' }
        elseif ($type -eq $want) { Add-Result $src 'Type' 'Pass' $want $type '' }
        else { Add-Result $src 'Type' 'Warn' $want $type 'Different from the source. Fine if that was the plan' }
    }

    $isMailbox = $type -match '^(User|Shared|Room|Equipment)Mailbox$'
    $mbx = if ($isMailbox) { Get-EXOMailbox -Identity $id -Properties ArchiveStatus, UserPrincipalName -ErrorAction SilentlyContinue }

    # ---------------------------------------------------- licence
    if ($p -and $p.Licence -and $p.Licence -ne 'None') {
        $u = Get-MgUser -UserId $dest.ExternalDirectoryObjectId -Property Id, AssignedLicenses -ErrorAction SilentlyContinue
        $names = @($u.AssignedLicenses | ForEach-Object { $skus["$($_.SkuId)"] } | Where-Object { $_ })
        if ($names -contains $p.Licence) { Add-Result $src 'Licence' 'Pass' $p.Licence ($names -join ';') '' }
        elseif ($names) { Add-Result $src 'Licence' 'Warn' $p.Licence ($names -join ';') 'Licensed, but not with the planned SKU' }
        else { Add-Result $src 'Licence' 'Fail' $p.Licence 'none' 'No licence assigned' }
    }

    # ---------------------------------------------------- archive
    if ($m -and $m.ArchiveStatus -eq 'Active') {
        $arch = if ($mbx) { "$($mbx.ArchiveStatus)" } else { 'no mailbox' }
        if ($arch -eq 'Active') { Add-Result $src 'Archive' 'Pass' 'Active' $arch '' }
        else { Add-Result $src 'Archive' 'Fail' 'Active' $arch 'The source had an archive. Enable it before copying archive content' }
    }

    # ---------------------------------------------------- addresses and sign-in after cutover
    if ($Domain -and $Domain -contains (Get-DomainOf $src) -and $type -ne 'MailContact') {
        $primary = "$($dest.PrimarySmtpAddress)".ToLower()
        if ($p) {
            if ($primary -eq $src) { Add-Result $src 'Address' 'Pass' $src $primary 'Primary' }
            else { Add-Result $src 'Address' 'Fail' $src $primary 'Old address is not primary. Run part 7 Switch-TargetAddresses' }
            if ($mbx) {
                $upn = "$($mbx.UserPrincipalName)".ToLower()
                if ($upn -eq $src) { Add-Result $src 'SignIn' 'Pass' $src $upn '' }
                else { Add-Result $src 'SignIn' 'Fail' $src $upn 'Sign-in name not switched' }
            }
        }
        elseif ($have -contains "smtp:$src") { Add-Result $src 'Address' 'Pass' $src $primary 'Present as an address' }
        else { Add-Result $src 'Address' 'Fail' $src $primary 'Old address missing. Mail to it will bounce' }
    }

    # ---------------------------------------------------- X500
    if ($SourceSnapshot) {
        if (-not $snap -or -not $snap.LegacyExchangeDN) { Add-Result $src 'X500' 'Warn' 'X500 address' 'not checked' 'Not in the source snapshot' }
        elseif ($have -contains "x500:$($snap.LegacyExchangeDN)".ToLower()) { Add-Result $src 'X500' 'Pass' 'present' 'present' '' }
        else { Add-Result $src 'X500' 'Fail' "X500:$($snap.LegacyExchangeDN)" 'missing' 'Replies to old mail will bounce. Run part 7 Switch-TargetAddresses' }
    }

    # ---------------------------------------------------- item counts
    if ($m -and $mbx) {
        $stats = Get-EXOMailboxStatistics -Identity $id -ErrorAction SilentlyContinue
        Test-Count $src 'Items' $m.ItemCount $stats.ItemCount
        if ($m.ArchiveStatus -eq 'Active' -and "$($mbx.ArchiveStatus)" -eq 'Active') {
            $astats = Get-EXOMailboxStatistics -Identity $id -Archive -ErrorAction SilentlyContinue
            Test-Count $src 'ArchiveItems' $m.ArchiveItemCount $astats.ItemCount
        }
    }

    # ---------------------------------------------------- OneDrive
    if ($null -ne $oneDrives -and $p -and $p.Kind -eq 'User' -and $mbx) {
        $upn = "$($mbx.UserPrincipalName)".ToLower()
        $hit = $oneDrives | Where-Object { ("$($_.Owner)".ToLower() -eq $upn) -or "$($_.Url)".ToLower().EndsWith((Get-PersonalPath $upn)) } | Select-Object -First 1
        if ($hit) { Add-Result $src 'OneDrive' 'Pass' 'exists' $hit.Url '' }
        else { Add-Result $src 'OneDrive' 'Fail' 'exists' 'none' 'No OneDrive for this sign-in name. Run part 4' }
    }

    # ---------------------------------------------------- temporary password must be changed
    if ($p -and $p.Kind -eq 'User' -and $type -eq 'UserMailbox' -and $dest.ExternalDirectoryObjectId) {
        $pw = Get-MgUser -UserId $dest.ExternalDirectoryObjectId -Property Id, PasswordProfile, LastPasswordChangeDateTime -ErrorAction SilentlyContinue
        $force = $pw.PasswordProfile.ForceChangePasswordNextSignIn   # Graph returns empty, not False, when the flag is off
        $changed = $pw.LastPasswordChangeDateTime
        if ($null -eq $pw) { Add-Result $src 'Password' 'Warn' 'change at next sign-in' 'not read' 'Could not read the user from Graph' }
        elseif ($force -eq $true) { Add-Result $src 'Password' 'Pass' 'change at next sign-in' 'True' '' }
        elseif ($CutoverDate -and $changed -and ([datetime]$changed) -ge $CutoverDate) {
            Add-Result $src 'Password' 'Pass' 'change at next sign-in' 'False' "Informational: password already changed on $(([datetime]$changed).ToString('yyyy-MM-dd HH:mm')), after cutover"
        }
        else {
            $when = if ($changed) { ([datetime]$changed).ToString('yyyy-MM-dd HH:mm') } else { 'unknown' }
            Add-Result $src 'Password' 'Warn' 'change at next sign-in' 'False or empty' "Last password change $when. If that's the person changing it after cutover, fine (pass -CutoverDate). Otherwise the temporary password you handed out stays valid: reset it with part 7 Switch-TargetAddresses -ResetPasswords"
        }
    }
}
Write-Progress -Activity 'Validating' -Completed

# ------------------------------------------------------------ tenant-wide: Conditional Access and MFA
if (-not $SkipTenantChecks) {
    try {
        $policies = [System.Collections.Generic.List[object]]::new()
        $uri = 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies'
        do {
            $page = Invoke-MgGraphRequest GET $uri
            foreach ($cap in $page.value) { $policies.Add($cap) }
            $uri = $page.'@odata.nextLink'
        } while ($uri)
        $live = @($policies | Where-Object { $_.state -in 'enabled', 'enabledForReportingButNotEnforced' })

        foreach ($cap in $live) {
            $loc = $cap.conditions.locations
            if ($loc -and (@($loc.includeLocations).Count -or @($loc.excludeLocations).Count)) {
                $mode = if ($cap.state -eq 'enabled') { 'enforced' } else { 'report-only' }
                Add-Result "CA: $($cap.displayName)" 'CA-Location' 'Warn' 'awareness' "location condition, $mode" 'Test sign-ins from outside the allowed locations (travelling, VPN, cloud PC) fail with error 53003. Test from where the users really are, or add your test account to the policy exclusions for the night and remove it after'
            }
        }

        $mfaAll = @($live | Where-Object {
                $_.state -eq 'enabled' -and @($_.conditions.users.includeUsers) -contains 'All' -and
                ((@($_.grantControls.builtInControls) -contains 'mfa') -or $_.grantControls.authenticationStrength)
            })
        $secDefaults = $false
        try { $secDefaults = [bool](Invoke-MgGraphRequest GET 'https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy').isEnabled } catch {}
        if ($mfaAll) { Add-Result 'Tenant' 'MFA' 'Pass' 'MFA for all users' (($mfaAll | ForEach-Object { $_.displayName }) -join '; ') '' }
        elseif ($secDefaults) { Add-Result 'Tenant' 'MFA' 'Pass' 'MFA for all users' 'Security defaults on' '' }
        else { Add-Result 'Tenant' 'MFA' 'Warn' 'MFA for all users' 'no policy' 'Finding only: no enabled Conditional Access policy requires MFA for all users, and security defaults are off. Raise it with the business owner. This script changes nothing' }
    }
    catch { Add-Result 'Tenant' 'CA-Location' 'Warn' 'policies read' 'not read' "Couldn't read Conditional Access ($($_.Exception.Message)). Needs Policy.Read.All, and Entra ID P1 for Conditional Access" }
}

# ------------------------------------------------------------ plan rows nobody mapped
foreach ($k in $plan.Keys) {
    if (-not $seen.Contains($k)) { Add-Result $k 'Mapping' 'Warn' 'a mapping row' 'none' 'In TargetBuildPlan.csv but not in the mapping file, so nothing else was checked' }
}

# ------------------------------------------------------------ distribution list members
if ($MailConfigFolder) {
    Write-Host 'Distribution lists' -ForegroundColor Cyan
    $counts = @{}
    Import-Csv (Join-Path $MailConfigFolder 'DistributionGroupMembers.csv') | ForEach-Object { $counts[$_.Group.ToLower()] = 1 + [int]$counts[$_.Group.ToLower()] }
    $mapIndex = @{}; foreach ($w in $mapping) { $mapIndex[$w.Source] = $w.Destination }
    foreach ($g in Import-Csv (Join-Path $MailConfigFolder 'DistributionGroups.csv')) {
        $src = $g.PrimarySmtpAddress.ToLower()
        $list = Get-Recipient -Filter "CustomAttribute15 -eq 'T2T:$src'" -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $list -and $mapIndex[$src]) { $list = Get-Recipient -Identity $mapIndex[$src] -ErrorAction SilentlyContinue }
        if (-not $list) { Add-Result $src 'ListMembers' 'Fail' 'list exists' 'none' 'Not found by tag or mapping. Check the part 5 import report'; continue }
        $want = [int]$counts[$src]
        $got = @(Get-DistributionGroupMember -Identity "$($list.PrimarySmtpAddress)" -ResultSize Unlimited -ErrorAction SilentlyContinue).Count
        if ($got -eq $want) { Add-Result $src 'ListMembers' 'Pass' $want $got '' }
        else { Add-Result $src 'ListMembers' 'Warn' $want $got 'Member count differs. Members without a mapping are skipped by part 5; check its report' }
    }
}

# ------------------------------------------------------------ report
$file = ".\Validation-$tenantName-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
$results | Export-Csv $file -NoTypeInformation -WhatIf:$false
Write-Host ''
$results | Group-Object Check | Sort-Object Name | ForEach-Object {
    $g = $_.Group
    [pscustomobject]@{
        Check = $_.Name
        Pass  = @($g | Where-Object Result -eq 'Pass').Count
        Warn  = @($g | Where-Object Result -eq 'Warn').Count
        Fail  = @($g | Where-Object Result -eq 'Fail').Count
    }
} | Format-Table -AutoSize | Out-String -Width 120 | Write-Host
$fails = @($results | Where-Object Result -eq 'Fail').Count
$warns = @($results | Where-Object Result -eq 'Warn').Count
$colour = if ($fails) { 'Red' } elseif ($warns) { 'Yellow' } else { 'Green' }
Write-Host ("{0} checks: {1} pass, {2} warn, {3} fail." -f $results.Count, ($results.Count - $fails - $warns), $warns, $fails) -ForegroundColor $colour
Write-Host "Report: $file" -ForegroundColor Cyan
if (-not $Domain) { Write-Host 'No -Domain given, so addresses and sign-in names were not checked. Add it after cutover.' -ForegroundColor Yellow }
if ($null -eq $oneDrives) { Write-Host 'OneDrives were not checked. Add -SharePointAdminUrl to include them.' -ForegroundColor Yellow }
