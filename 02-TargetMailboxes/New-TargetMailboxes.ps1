<#
.SYNOPSIS
    Builds the destination mailboxes for a Microsoft 365 tenant-to-tenant migration from a source inventory.

.DESCRIPTION
    TL;DR: Two steps. -Plan reads 04-Mailboxes.csv from Invoke-SourceTenantInventory and writes TargetBuildPlan.csv
    for you to review. -Apply reads that plan and builds the destination: shared mailboxes, user mailboxes,
    licences, archives and auto-expanding archives. Re-run -Apply as often as you like. It only does what is
    still missing.

    The usual flow:
      1. -Apply -StageOnly                 Every mailbox as an unlicensed shared mailbox. Nobody can sign in, no
                                           licences needed, and you can start copying mail straight away.
      2. -Apply -SharedMailboxesOnly       Licences arrive: license and archive the shared mailboxes. People stay
                                           staged until cutover.
      3. -Apply                            Cutover: convert people to user mailboxes, enable sign-in, set temporary
                                           passwords, license, enable archives. Run it twice, about 15 minutes
                                           apart, so the archives and auto-expanding archives finish.

    Everyone is created on -TargetDomain (normally the destination's onmicrosoft.com domain). Use
    -StagingSmtpDomain to give staged mailboxes a friendlier primary address on a domain the destination already
    owns. The source domains are added after they move.

.PARAMETER Plan
    Build TargetBuildPlan.csv from the inventory. Changes nothing in any tenant.

.PARAMETER Apply
    Build the destination from TargetBuildPlan.csv. Supports -WhatIf.

.PARAMETER TargetDomain
    The destination domain every mailbox is created on, e.g. contoso.onmicrosoft.com. -Apply refuses to run if the
    signed-in tenant doesn't own it.

.PARAMETER InventoryFolder
    Inventory folder containing 04-Mailboxes.csv. The newest Inventory-* folder here is used if omitted.

.PARAMETER SourceDomain
    The source tenant's main mail domain. Mailboxes on it keep their alias as-is; mailboxes on other domains get
    alias-<domainname> so they can't clash. Defaults to the most common domain in the inventory.

.PARAMETER Exclude
    Source addresses or aliases to leave out (admin accounts, people who already exist in the destination).

.PARAMETER TreatAsShared
    Source user mailboxes to build as shared mailboxes, typically leavers whose mail you want to keep.

.PARAMETER UserLicenceSku
    SkuPartNumber for people. Default SPB (Microsoft 365 Business Premium).

.PARAMETER SharedLicenceSku
    SkuPartNumber for shared mailboxes that need one (archive, or over the free size). Default EXCHANGEENTERPRISE
    (Exchange Online Plan 2).

.PARAMETER SharedQuotaGB
    Size an unlicensed shared mailbox can hold. Default 50. Bigger mailboxes are planned with -SharedLicenceSku.

.PARAMETER AutoExpandThresholdGB
    Archives at or above this size get an auto-expanding archive. Default 90.

.PARAMETER DisplayNameSuffix
    Optional regex stripped from display names, e.g. '\s+-\s+Contoso$'.

.PARAMETER UsageLocation
    Two-letter country code, required for licensing. Required with -Apply.

.PARAMETER StageOnly
    With -Apply: create every planned mailbox as an unlicensed shared mailbox. No licences, no archives.

.PARAMETER SharedMailboxesOnly
    With -Apply: license and archive the Kind 'Shared' rows only. People stay staged.

.PARAMETER StagingSmtpDomain
    With -Apply: set the primary SMTP address of mailboxes from -SourceDomain to alias@StagingSmtpDomain. The
    TargetDomain address stays as a proxy, so mappings keep working.

.PARAMETER ResetPasswords
    With -Apply: give every person a new temporary password, even if their account is already enabled. Useful at
    cutover, when you hand passwords out.

.EXAMPLE
    .\New-TargetMailboxes.ps1 -Plan -TargetDomain contoso.onmicrosoft.com -Exclude admin@fabrikam.com

.EXAMPLE
    .\New-TargetMailboxes.ps1 -Apply -TargetDomain contoso.onmicrosoft.com -UsageLocation GB -StageOnly -WhatIf

.EXAMPLE
    .\New-TargetMailboxes.ps1 -Apply -TargetDomain contoso.onmicrosoft.com -UsageLocation GB

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 2.0.0
    Date    : 2026-11-07
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: ExchangeOnlineManagement 3.x, Microsoft.Graph.Authentication, Microsoft.Graph.Users,
              Microsoft.Graph.Identity.DirectoryManagement. Run -Apply as a Global Admin of the DESTINATION.
    The passwords file written by -Apply is sensitive. Move it somewhere secure and delete the local copy.
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Plan')]
param(
    [Parameter(ParameterSetName = 'Plan', Mandatory)][switch]$Plan,
    [Parameter(ParameterSetName = 'Apply', Mandatory)][switch]$Apply,
    [Parameter(Mandatory)][string]$TargetDomain,
    [string]$PlanFile = '.\TargetBuildPlan.csv',

    [Parameter(ParameterSetName = 'Plan')][string]$InventoryFolder,
    [Parameter(ParameterSetName = 'Plan')][string]$SourceDomain,
    [Parameter(ParameterSetName = 'Plan')][string[]]$Exclude = @(),
    [Parameter(ParameterSetName = 'Plan')][string[]]$TreatAsShared = @(),
    [Parameter(ParameterSetName = 'Plan')][string]$UserLicenceSku = 'SPB',
    [Parameter(ParameterSetName = 'Plan')][string]$SharedLicenceSku = 'EXCHANGEENTERPRISE',
    [Parameter(ParameterSetName = 'Plan')][double]$SharedQuotaGB = 50,
    [Parameter(ParameterSetName = 'Plan')][double]$AutoExpandThresholdGB = 90,
    [Parameter(ParameterSetName = 'Plan')][string]$DisplayNameSuffix,

    [Parameter(ParameterSetName = 'Apply', Mandatory)][ValidatePattern('^[A-Za-z]{2}$')][string]$UsageLocation,
    [Parameter(ParameterSetName = 'Apply')][switch]$StageOnly,
    [Parameter(ParameterSetName = 'Apply')][switch]$SharedMailboxesOnly,
    [Parameter(ParameterSetName = 'Apply')][string]$StagingSmtpDomain,
    [Parameter(ParameterSetName = 'Apply')][switch]$ResetPasswords
)

$ErrorActionPreference = 'Stop'

function New-RandomPassword {
    $rng = [System.Security.Cryptography.RandomNumberGenerator]
    $sets = 'ABCDEFGHJKLMNPQRSTUVWXYZ', 'abcdefghijkmnpqrstuvwxyz', '23456789', '!@#$%*?'
    $all = -join $sets
    $chars = [System.Collections.Generic.List[char]]::new()
    foreach ($s in $sets) { $chars.Add($s[$rng::GetInt32($s.Length)]) }          # one of each class
    1..12 | ForEach-Object { $chars.Add($all[$rng::GetInt32($all.Length)]) }
    return -join ($chars | Sort-Object { $rng::GetInt32([int]::MaxValue) })      # shuffle
}

function Get-ErrorText($err) {
    if ($err.ErrorDetails.Message) { $err.ErrorDetails.Message } else { $err.Exception.Message }
}

# ============================================================ PLAN
if ($Plan) {
    if (-not $InventoryFolder) {
        $InventoryFolder = (Get-ChildItem -Directory -Filter 'Inventory-*' |
            Where-Object { Test-Path (Join-Path $_.FullName '04-Mailboxes.csv') } |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
    }
    if (-not $InventoryFolder -or -not (Test-Path (Join-Path $InventoryFolder '04-Mailboxes.csv'))) {
        throw 'No inventory folder with 04-Mailboxes.csv found. Pass -InventoryFolder.'
    }
    Write-Host "Using inventory: $InventoryFolder" -ForegroundColor Cyan

    $mailboxes = Import-Csv (Join-Path $InventoryFolder '04-Mailboxes.csv') |
        Where-Object { $_.RecipientTypeDetails -ne 'DiscoveryMailbox' }
    if (-not $SourceDomain) {
        $SourceDomain = ($mailboxes | ForEach-Object { $_.PrimarySmtpAddress.Split('@')[1].ToLower() } |
            Group-Object | Sort-Object Count -Descending | Select-Object -First 1).Name
    }
    Write-Host "Main source domain: $SourceDomain" -ForegroundColor Cyan

    $excludeSet = @($Exclude | ForEach-Object { $_.ToLower() })
    $sharedSet  = @($TreatAsShared | ForEach-Object { $_.ToLower() })
    $used = [System.Collections.Generic.HashSet[string]]::new()

    $rows = foreach ($m in $mailboxes) {
        $address = $m.PrimarySmtpAddress.ToLower()
        $local, $domain = $address.Split('@')
        $main = $domain -eq $SourceDomain
        $alias = if ($main) { $local } else { "$local-" + $domain.Split('.')[0] }
        while (-not $used.Add($alias)) { $alias = "$alias-1" }

        $sizeGB    = [double]('0' + $m.SizeGB)
        $archiveGB = [double]('0' + $m.ArchiveSizeGB)
        $hasArchive = $m.ArchiveStatus -eq 'Active'
        $include = $true; $note = ''

        switch -Regex ($m.RecipientTypeDetails) {
            '^UserMailbox$' {
                if ($sharedSet -contains $address -or ($main -and $sharedSet -contains $local)) { $kind = 'Shared' }
                else { $kind = 'User' }
            }
            '^(RoomMailbox|EquipmentMailbox|SchedulingMailbox)$' {
                $kind = 'Shared'; $include = $false; $note = "$($m.RecipientTypeDetails): build by hand, can't be converted"
            }
            default { $kind = 'Shared' }
        }
        if ($excludeSet -contains $address -or ($main -and $excludeSet -contains $local)) {
            $include = $false; $note = 'Excluded by -Exclude'
        }

        $licence = if ($kind -eq 'User') { $UserLicenceSku }
                   elseif ($hasArchive -or $sizeGB -ge ($SharedQuotaGB * 0.9)) { $SharedLicenceSku }
                   else { 'None' }

        $name = $m.DisplayName -replace '\s{2,}', ' '
        if ($DisplayNameSuffix) { $name = $name -replace $DisplayNameSuffix, '' }

        [pscustomobject]@{
            Include         = $include
            SourceAddress   = $m.PrimarySmtpAddress
            MainDomain      = $main
            DisplayName     = $name.Trim()
            Alias           = $alias
            TargetAddress   = "$alias@$TargetDomain"
            Kind            = $kind
            Licence         = $licence
            Archive         = $hasArchive
            AutoExpanding   = ($hasArchive -and $archiveGB -ge $AutoExpandThresholdGB)
            SourceType      = $m.RecipientTypeDetails
            SourceMailGB    = $m.SizeGB
            SourceArchiveGB = $m.ArchiveSizeGB
            SourceAliases   = $m.Aliases
            Note            = $note
        }
    }

    $rows | Sort-Object Kind, Licence, DisplayName | Export-Csv $PlanFile -NoTypeInformation
    $inc = @($rows | Where-Object Include)
    Write-Host ''
    $inc | Group-Object Kind, Licence | Select-Object @{ n = 'Kind, Licence'; e = { $_.Name } }, Count | Format-Table -AutoSize | Out-Host
    Write-Host ("Plan written to {0}: {1} included, {2} excluded." -f (Resolve-Path $PlanFile), $inc.Count,
        @($rows | Where-Object { -not $_.Include }).Count) -ForegroundColor Green
    Write-Host 'Review it before -Apply. Exclude anyone who already has a mailbox in the destination, and rename any alias that already exists there.' -ForegroundColor Yellow
    return
}

# ============================================================ APPLY
if ($StageOnly -and $SharedMailboxesOnly) { throw 'Use -StageOnly or -SharedMailboxesOnly, not both.' }
if (-not (Test-Path $PlanFile)) { throw "No $PlanFile. Run -Plan first." }
$planRows = @(Import-Csv $PlanFile | Where-Object { $_.Include -eq 'True' })

Connect-MgGraph -TenantId $TargetDomain -NoWelcome -Scopes 'User.ReadWrite.All', 'User-PasswordProfile.ReadWrite.All',
    'Organization.Read.All', 'Directory.Read.All'
Connect-ExchangeOnline -ShowBanner:$false

# Safety: refuse to run against the wrong tenant.
$graphDomains = @((Get-MgDomain).Id)
$exoDomains   = @((Get-AcceptedDomain).DomainName)
if (($graphDomains -notcontains $TargetDomain) -or ($exoDomains -notcontains $TargetDomain)) {
    throw "The signed-in tenant doesn't own $TargetDomain. Disconnect (Disconnect-MgGraph; Disconnect-ExchangeOnline) and sign in to the destination."
}
if ($StagingSmtpDomain -and $exoDomains -notcontains $StagingSmtpDomain) {
    throw "$StagingSmtpDomain isn't an accepted domain in this tenant."
}
Write-Host "Destination confirmed: $TargetDomain" -ForegroundColor Cyan

$seats = @{}
foreach ($s in Get-MgSubscribedSku -All) {
    $seats[$s.SkuPartNumber] = [pscustomobject]@{ SkuId = $s.SkuId; Free = [int]$s.PrepaidUnits.Enabled - [int]$s.ConsumedUnits }
}

$status = @(); $passwords = @()
foreach ($r in $planRows) {
    $state = [ordered]@{ TargetAddress = $r.TargetAddress; Kind = $r.Kind; Created = ''; Licence = ''; Archive = ''; AutoExpanding = ''; Note = '' }
    $userId = $null
    $buildAsUser = $r.Kind -eq 'User' -and -not $StageOnly -and -not $SharedMailboxesOnly

    try {
        # ---------------- create or convert
        $existing = Get-EXOMailbox -Identity $r.TargetAddress -Properties RecipientTypeDetails, ExternalDirectoryObjectId -ErrorAction SilentlyContinue

        if ($buildAsUser -and $existing) {
            # Staged earlier as a shared mailbox, or converted by an earlier run that stopped part-way.
            $userId = $existing.ExternalDirectoryObjectId
            if ($existing.RecipientTypeDetails -eq 'SharedMailbox') {
                if ($PSCmdlet.ShouldProcess($r.TargetAddress, 'Convert shared mailbox to user mailbox')) {
                    Set-Mailbox -Identity $userId -Type Regular
                    $state.Created = 'Converted to user'
                }
            }
            else { $state.Created = 'Exists' }

            $u = Get-MgUser -UserId $userId -Property Id, AccountEnabled, UsageLocation
            if (-not $u.UsageLocation -and $PSCmdlet.ShouldProcess($r.TargetAddress, "Set usage location $UsageLocation")) {
                Update-MgUser -UserId $userId -UsageLocation $UsageLocation
            }
            if ((-not $u.AccountEnabled -or $ResetPasswords) -and
                $PSCmdlet.ShouldProcess($r.TargetAddress, 'Enable sign-in and set a temporary password')) {
                try {
                    $pw = New-RandomPassword
                    Update-MgUser -UserId $userId -AccountEnabled:$true -PasswordProfile @{ Password = $pw; ForceChangePasswordNextSignIn = $true }
                    $passwords += [pscustomobject]@{ DisplayName = $r.DisplayName; SignIn = $r.TargetAddress; TempPassword = $pw; SourceAddress = $r.SourceAddress }
                    $state.Created += ', password set'
                }
                catch {
                    # Carry on so the licence and archive still happen. Re-run to retry the password.
                    $state.Note = 'Password not set: ' + (Get-ErrorText $_)
                    Write-Warning "$($r.TargetAddress): $($state.Note)"
                }
            }
        }
        elseif ($buildAsUser) {
            $u = Get-MgUser -Filter "userPrincipalName eq '$($r.TargetAddress)'" -Property Id -ErrorAction SilentlyContinue
            if (-not $u -and $PSCmdlet.ShouldProcess($r.TargetAddress, 'Create user')) {
                $pw = New-RandomPassword
                $u = New-MgUser -DisplayName $r.DisplayName -MailNickname $r.Alias -UserPrincipalName $r.TargetAddress `
                    -AccountEnabled -UsageLocation $UsageLocation -PasswordProfile @{ Password = $pw; ForceChangePasswordNextSignIn = $true }
                $passwords += [pscustomobject]@{ DisplayName = $r.DisplayName; SignIn = $r.TargetAddress; TempPassword = $pw; SourceAddress = $r.SourceAddress }
                $state.Created = 'Created'
            }
            elseif ($u) { $state.Created = 'Exists' }
            $userId = $u.Id
        }
        elseif ($r.Kind -eq 'User' -and $SharedMailboxesOnly) {
            $state.Created = if ($existing) { 'Staged' } else { 'Missing' }
            $state.Licence = 'Deferred to cutover'
            $status += [pscustomobject]$state
            continue
        }
        else {
            # Shared mailbox, or anyone during -StageOnly.
            $mbx = $existing
            if (-not $mbx -and $PSCmdlet.ShouldProcess($r.TargetAddress, 'Create shared mailbox')) {
                $mbx = New-Mailbox -Shared -Name $r.Alias -Alias $r.Alias -DisplayName $r.DisplayName -PrimarySmtpAddress $r.TargetAddress -ErrorAction Stop
                if (-not $mbx) { throw "New-Mailbox returned nothing for $($r.TargetAddress). Is the alias or name already in use?" }
                $state.Created = 'Created'
            }
            elseif ($mbx) { $state.Created = 'Exists' }
            if ($mbx) { $userId = $mbx.ExternalDirectoryObjectId }
            if ($mbx -and -not $userId) {
                # A brand-new shared mailbox can take a minute to appear in Entra ID.
                $state.Licence = 'PENDING: directory object not ready, re-run shortly'
                $status += [pscustomobject]$state
                continue
            }
        }

        # ---------------- friendlier primary address while staged
        if ($StagingSmtpDomain -and $r.MainDomain -eq 'True' -and $userId) {
            $want = ($r.SourceAddress.Split('@')[0]).ToLower() + "@$StagingSmtpDomain"
            $now = (Get-EXOMailbox -Identity $userId).PrimarySmtpAddress
            if ($now -ne $want -and $PSCmdlet.ShouldProcess($r.TargetAddress, "Set primary address $want")) {
                Set-Mailbox -Identity $userId -WindowsEmailAddress $want
            }
        }

        if ($StageOnly) {
            $state.Licence = 'Deferred (StageOnly)'
            $status += [pscustomobject]$state
            continue
        }

        # ---------------- licence
        if ($r.Licence -ne 'None' -and $userId) {
            $u = Get-MgUser -UserId $userId -Property Id, UsageLocation, AssignedLicenses
            $sku = $seats[$r.Licence]
            if (-not $sku) { $state.Licence = "SKU $($r.Licence) not in tenant" }
            elseif ($u.AssignedLicenses.SkuId -contains $sku.SkuId) { $state.Licence = 'Assigned' }
            elseif ($sku.Free -lt 1) { $state.Licence = "PENDING: no free $($r.Licence)" }
            elseif ($PSCmdlet.ShouldProcess($r.TargetAddress, "Assign $($r.Licence)")) {
                if (-not $u.UsageLocation) { Update-MgUser -UserId $userId -UsageLocation $UsageLocation }
                # Set-MgUserLicense drops an empty -RemoveLicenses array and Graph returns 400, so call the API directly.
                $body = @{ addLicenses = @(@{ skuId = $sku.SkuId; disabledPlans = @() }); removeLicenses = @() } | ConvertTo-Json -Depth 5
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$userId/assignLicense" `
                    -Body $body -ContentType 'application/json' | Out-Null
                $sku.Free--
                $state.Licence = 'Assigned now'
            }
        }
        elseif ($r.Licence -eq 'None') { $state.Licence = 'Not needed' }

        # ---------------- archive and auto-expanding archive
        if ($r.Archive -eq 'True' -and $userId) {
            $m = Get-Mailbox -Identity $userId -ErrorAction SilentlyContinue
            if (-not $m) { $state.Archive = 'PENDING: mailbox provisioning, re-run in 15 min' }
            elseif ($state.Licence -notin 'Assigned', 'Assigned now') { $state.Archive = 'PENDING: needs licence first' }
            else {
                if ($m.ArchiveStatus -ne 'Active') {
                    if ($PSCmdlet.ShouldProcess($r.TargetAddress, 'Enable archive')) {
                        Enable-Mailbox -Identity $userId -Archive | Out-Null
                        $state.Archive = 'Enabled now'
                    }
                }
                else { $state.Archive = 'Active' }

                if ($r.AutoExpanding -eq 'True') {
                    $m = Get-Mailbox -Identity $userId
                    if ($m.AutoExpandingArchiveEnabled) { $state.AutoExpanding = 'Enabled' }
                    elseif ($m.ArchiveStatus -ne 'Active') { $state.AutoExpanding = 'PENDING: re-run after archive is active' }
                    elseif ($PSCmdlet.ShouldProcess($r.TargetAddress, 'Enable auto-expanding archive')) {
                        Enable-Mailbox -Identity $userId -AutoExpandingArchive | Out-Null
                        $state.AutoExpanding = 'Enabled now'
                    }
                }
            }
        }
    }
    catch {
        $state.Note = Get-ErrorText $_
        Write-Warning "$($r.TargetAddress): $($state.Note)"
    }
    $status += [pscustomobject]$state
}

$stamp = Get-Date -Format 'yyyyMMdd-HHmm'
$status | Export-Csv ".\TargetBuild-Status-$stamp.csv" -NoTypeInformation
$status | Format-Table TargetAddress, Kind, Created, Licence, Archive, AutoExpanding -AutoSize | Out-Host
if ($passwords) {
    $passwords | Export-Csv ".\TargetBuild-Passwords-$stamp.csv" -NoTypeInformation
    Write-Warning "Temporary passwords saved to TargetBuild-Passwords-$stamp.csv. Move it somewhere secure and delete this copy."
}
$pending = @($status | Where-Object { ($_.Licence, $_.Archive, $_.AutoExpanding) -match 'PENDING' })
$failed  = @($status | Where-Object { $_.Note -and $_.Note -notlike 'Password not set*' })
if ($failed)  { Write-Host "$($failed.Count) mailbox(es) hit an error. See the Note column in TargetBuild-Status-$stamp.csv." -ForegroundColor Red }
if ($pending) { Write-Host "$($pending.Count) item(s) pending. Re-run -Apply once licences are free or in about 15 minutes." -ForegroundColor Yellow }
elseif (-not $failed) { Write-Host 'All planned mailboxes are in place.' -ForegroundColor Green }
