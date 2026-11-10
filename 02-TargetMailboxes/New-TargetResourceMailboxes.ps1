<#
.SYNOPSIS
    Rebuilds room and equipment mailboxes in the destination tenant, with their capacity and booking rules, from the
    source inventory or a live read of the source.

.DESCRIPTION
    TL;DR: Rooms and equipment can't be staged as shared mailboxes and converted later, so New-TargetMailboxes leaves
    them out. This script builds them properly. It reads 15-ResourceMailboxes.csv from Invoke-SourceTenantInventory
    (or reads the source live with -SourceAdminUpn), creates each one in the destination with New-Mailbox -Room or
    -Equipment on a staging address, copies capacity, and sets calendar processing: AutomateProcessing,
    BookingWindowInDays, MaximumDurationInMinutes, AllowConflicts, AllowRecurringMeetings, ResourceDelegates and
    BookInPolicy / RequestInPolicy (people mapped through the mapping CSV), plus the subject and privacy settings.
    Safe to re-run: every mailbox it creates is tagged CustomAttribute15 = 'T2T:<source address>' and recognised
    next time. Supports -WhatIf. Writes TargetResourceMailboxes-<date>.csv.

    Tip: run it after 02b has built UserMapping.csv, so delegates and booking policies map to real people. Then run
    02b again so the rooms themselves appear in the mapping (it matches on the CustomAttribute15 tag).
    Trap: a delegate or policy entry with no mapping is left out and reported, not guessed. Read the Detail column.
    Trap: room and equipment mailboxes need no licence. Don't give them one by habit.

.PARAMETER TargetDomain
    The destination domain the sign-in name is created on, e.g. contoso.onmicrosoft.com. The script refuses to run
    if the signed-in tenant doesn't own it.

.PARAMETER InventoryFolder
    Inventory folder containing 15-ResourceMailboxes.csv (Invoke-SourceTenantInventory 1.1.0 or later). The newest
    Inventory-* folder here that has the file is used if omitted.

.PARAMETER SourceAdminUpn
    Instead of the inventory: read rooms and equipment live from the source, signed in as this account. The script
    signs in to the source, reads, saves a snapshot CSV, disconnects, then signs in to the destination.

.PARAMETER MappingFile
    CSV with Source and Destination columns (UserMapping.csv from 02b). Used for ResourceDelegates, BookInPolicy and
    RequestInPolicy. Without it, those lists are left empty and reported.

.PARAMETER StagingSmtpDomain
    Optional. Primary SMTP address for the new mailboxes, e.g. contoso.com, until the source domain moves at
    cutover. Defaults to -TargetDomain.

.PARAMETER AliasPrefix
    Optional prefix for new aliases, for when the destination already has a 'boardroom'. Example: 'fab-'.

.PARAMETER Exclude
    Source addresses or aliases to leave out.

.PARAMETER AdminUpn
    Optional destination admin sign-in name, e.g. admin@contoso.onmicrosoft.com. Any Exchange Online session for
    another account or tenant is disconnected first, and the script prints who and where before changing anything.

.PARAMETER TenantId
    Optional destination tenant ID (GUID). The script stops if Exchange Online is signed in anywhere else.

.EXAMPLE
    .\New-TargetResourceMailboxes.ps1 -TargetDomain contoso.onmicrosoft.com -MappingFile .\UserMapping.csv -WhatIf

.EXAMPLE
    .\New-TargetResourceMailboxes.ps1 -TargetDomain contoso.onmicrosoft.com -StagingSmtpDomain contoso.com `
        -MappingFile .\UserMapping.csv -AdminUpn admin@contoso.onmicrosoft.com

.EXAMPLE
    .\New-TargetResourceMailboxes.ps1 -SourceAdminUpn admin@fabrikam.onmicrosoft.com -TargetDomain contoso.onmicrosoft.com `
        -MappingFile .\UserMapping.csv -AdminUpn admin@contoso.onmicrosoft.com -AliasPrefix fab-

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.0.0
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: ExchangeOnlineManagement 3.x. Exchange Administrator (or Recipient Administrator) in the destination;
              read access in the source if you use -SourceAdminUpn.
    Calendar items (existing bookings) are content: your copy tool moves those. This script builds the mailbox and
    its rules, so the room keeps taking bookings the way it did.
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Inventory')]
param(
    [Parameter(Mandatory)][string]$TargetDomain,
    [Parameter(ParameterSetName = 'Inventory')][string]$InventoryFolder,
    [Parameter(ParameterSetName = 'Live', Mandatory)][string]$SourceAdminUpn,
    [string]$MappingFile,
    [string]$StagingSmtpDomain,
    [string]$AliasPrefix = '',
    [string[]]$Exclude = @(),
    [string]$AdminUpn,
    [ValidatePattern('^[0-9a-fA-F-]{36}$')][string]$TenantId
)

$ErrorActionPreference = 'Stop'
if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) { throw 'Install-Module ExchangeOnlineManagement -Scope CurrentUser first.' }

function Get-Tag([string]$Source) { "T2T:$($Source.ToLower())" }
function ConvertTo-Bool($v) { if ("$v" -match '^(?i)(true|false)$') { [bool]::Parse("$v") } else { $null } }

function Connect-Exo([string]$Upn, [string]$Tenant, [string]$Label) {
    # Close any session for another account or tenant first, so Windows can't quietly reuse it.
    $WhatIfPreference = $false
    $c = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object State -eq 'Connected' | Select-Object -First 1
    if ($c -and (($Upn -and "$($c.UserPrincipalName)" -ne $Upn) -or ($Tenant -and "$($c.TenantID)" -ne $Tenant))) {
        Write-Host "Exchange Online is signed in as $($c.UserPrincipalName) to $($c.TenantID). Reconnecting for the $Label." -ForegroundColor Yellow
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        $c = $null
    }
    if (-not $c) {
        $p = @{ ShowBanner = $false }
        if ($Upn) { $p.UserPrincipalName = $Upn }
        Connect-ExchangeOnline @p
    }
    $c = Get-ConnectionInformation | Where-Object State -eq 'Connected' | Select-Object -First 1
    if ($Upn -and "$($c.UserPrincipalName)" -ne $Upn) { throw "Exchange Online signed in as $($c.UserPrincipalName), not $Upn. Pick the right account, or use a private browser window." }
    if ($Tenant -and "$($c.TenantID)" -ne $Tenant) { throw "Exchange Online is signed in to tenant $($c.TenantID), not $Tenant." }
    Write-Host "Signed in as $($c.UserPrincipalName) to tenant $($c.TenantID) ($Label)" -ForegroundColor Cyan
    $c
}

# ------------------------------------------------------------ read the source
$sourceTenant = $null
if ($PSCmdlet.ParameterSetName -eq 'Live') {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    $sc = Connect-Exo $SourceAdminUpn $null 'source'
    $sourceTenant = "$($sc.TenantID)"
    $cache = @{}
    function Resolve-Smtp($id) {
        if (-not $id) { return $null }
        $k = "$id"
        if (-not $cache.ContainsKey($k)) {
            $r = Get-EXORecipient -Identity $k -ErrorAction SilentlyContinue
            $cache[$k] = if ($r) { "$($r.PrimarySmtpAddress)".ToLower() } else { "unresolved:$k" }
        }
        $cache[$k]
    }
    $rows = foreach ($m in @(Get-EXOMailbox -RecipientTypeDetails RoomMailbox, EquipmentMailbox -ResultSize Unlimited)) {
        $id = $m.ExternalDirectoryObjectId
        $full = Get-Mailbox -Identity $id
        $cp = Get-CalendarProcessing -Identity $id
        [pscustomobject]@{
            DisplayName                    = $m.DisplayName
            PrimarySmtpAddress             = "$($m.PrimarySmtpAddress)".ToLower()
            Alias                          = $full.Alias
            RecipientTypeDetails           = $m.RecipientTypeDetails
            ResourceCapacity               = $full.ResourceCapacity
            Office                         = $full.Office
            HiddenFromAddressListsEnabled  = $full.HiddenFromAddressListsEnabled
            AutomateProcessing             = $cp.AutomateProcessing
            BookingWindowInDays            = $cp.BookingWindowInDays
            MaximumDurationInMinutes       = $cp.MaximumDurationInMinutes
            AllowConflicts                 = $cp.AllowConflicts
            AllowRecurringMeetings         = $cp.AllowRecurringMeetings
            AllBookInPolicy                = $cp.AllBookInPolicy
            AllRequestInPolicy             = $cp.AllRequestInPolicy
            BookInPolicy                   = (@($cp.BookInPolicy | ForEach-Object { Resolve-Smtp $_ }) -join ';')
            RequestInPolicy                = (@($cp.RequestInPolicy | ForEach-Object { Resolve-Smtp $_ }) -join ';')
            ResourceDelegates              = (@($cp.ResourceDelegates | ForEach-Object { Resolve-Smtp $_ }) -join ';')
            AddOrganizerToSubject          = $cp.AddOrganizerToSubject
            DeleteSubject                  = $cp.DeleteSubject
            DeleteComments                 = $cp.DeleteComments
            RemovePrivateProperty          = $cp.RemovePrivateProperty
            ProcessExternalMeetingMessages = $cp.ProcessExternalMeetingMessages
        }
    }
    $rows = @($rows)
    $snap = ".\ResourceMailboxes-Source-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
    $rows | Export-Csv $snap -NoTypeInformation -WhatIf:$false
    Write-Host "$($rows.Count) room and equipment mailbox(es) read from the source. Snapshot: $snap" -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue -WhatIf:$false | Out-Null
}
else {
    if (-not $InventoryFolder) {
        $InventoryFolder = (Get-ChildItem -Directory -Filter 'Inventory-*' -ErrorAction SilentlyContinue |
            Where-Object { Test-Path (Join-Path $_.FullName '15-ResourceMailboxes.csv') } |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
    }
    if (-not $InventoryFolder -or -not (Test-Path (Join-Path $InventoryFolder '15-ResourceMailboxes.csv'))) {
        throw 'No inventory folder with 15-ResourceMailboxes.csv found. Re-run Invoke-SourceTenantInventory 1.1.0 or later, pass -InventoryFolder, or read the source live with -SourceAdminUpn.'
    }
    Write-Host "Using inventory: $InventoryFolder" -ForegroundColor Cyan
    $rows = @(Import-Csv (Join-Path $InventoryFolder '15-ResourceMailboxes.csv'))
}
if (-not $rows) { Write-Host 'No room or equipment mailboxes to build.' -ForegroundColor Green; return }

# ------------------------------------------------------------ mapping
$map = @{}
if ($MappingFile) {
    if (-not (Test-Path $MappingFile)) { throw "Mapping file $MappingFile not found." }
    Import-Csv $MappingFile | Where-Object { $_.Source -and $_.Destination } | ForEach-Object { $map[$_.Source.Trim().ToLower()] = $_.Destination.Trim() }
    Write-Host "$($map.Count) mappings loaded." -ForegroundColor Cyan
}
else { Write-Warning 'No -MappingFile. Delegates and booking policies will be left empty and reported.' }

# ------------------------------------------------------------ destination
$dc = Connect-Exo $AdminUpn $TenantId 'destination'
if ($sourceTenant -and "$($dc.TenantID)" -eq $sourceTenant) { throw 'The destination session is the same tenant as the source. Sign in to the destination with -AdminUpn.' }
$accepted = @((Get-AcceptedDomain).DomainName | ForEach-Object { "$_".ToLower() })
if ($accepted -notcontains $TargetDomain.ToLower()) { throw "The signed-in tenant doesn't own $TargetDomain. Sign in to the destination (use -AdminUpn)." }
if ($StagingSmtpDomain -and $accepted -notcontains $StagingSmtpDomain.ToLower()) { throw "$StagingSmtpDomain isn't an accepted domain in this tenant." }
$addressDomain = if ($StagingSmtpDomain) { $StagingSmtpDomain.ToLower() } else { $TargetDomain.ToLower() }
Write-Host "Destination confirmed: $TargetDomain. New addresses on $addressDomain" -ForegroundColor Cyan

$log = [System.Collections.Generic.List[object]]::new()
function Write-Log($Area, $Source, $Destination, $Status, $Detail) {
    $log.Add([pscustomobject]@{ Area = $Area; Source = $Source; Destination = $Destination; Status = $Status; Detail = $Detail })
    $c = @{ Created = 'Green'; Set = 'Green'; Exists = 'Gray'; Skipped = 'Yellow'; Pending = 'Yellow'; Problem = 'Red'; WhatIf = 'Cyan' }[$Status]
    Write-Host ("  [{0}] {1}: {2} {3}" -f $Status, $Area, $Source, $Detail) -ForegroundColor $c
}
function Get-MappedList([string]$List, [ref]$Missing) {
    foreach ($a in @("$List" -split ';' | Where-Object { $_ })) {
        $to = $map[$a.Trim().ToLower()]
        if ($to) { $to }
        elseif ($a -like 'unresolved:*') { $Missing.Value += "$($a -replace '^unresolved:', '') (no longer in the source)" }
        else { $Missing.Value += $a }
    }
}

$excludeSet = @($Exclude | ForEach-Object { $_.ToLower() })
foreach ($r in $rows) {
    $src = "$($r.PrimarySmtpAddress)".ToLower()
    $local = $src.Split('@')[0]
    if ($excludeSet -contains $src -or $excludeSet -contains $local -or $excludeSet -contains "$($r.Alias)".ToLower()) {
        Write-Log 'Mailbox' $src '' 'Skipped' 'Excluded by -Exclude'; continue
    }
    $isRoom = $r.RecipientTypeDetails -eq 'RoomMailbox'
    $alias = ("$AliasPrefix$local" -replace '[^a-zA-Z0-9._-]', '').ToLower()
    $signIn = "$alias@$($TargetDomain.ToLower())"
    $primary = "$alias@$addressDomain"
    $id = $null

    # ---------------- create, or find the one an earlier run made
    $ours = Get-Recipient -Filter "CustomAttribute15 -eq '$(Get-Tag $src)'" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($ours) {
        $id = "$($ours.PrimarySmtpAddress)"
        $map[$src] = $id
        Write-Log 'Mailbox' $src $id 'Exists' 'Created by an earlier run'
    }
    elseif ((Get-Recipient -Identity $signIn -ErrorAction SilentlyContinue) -or (Get-Recipient -Identity $primary -ErrorAction SilentlyContinue) -or
            (Get-Recipient -Identity $alias -ErrorAction SilentlyContinue)) {
        Write-Log 'Mailbox' $src $primary 'Skipped' "Alias or address '$alias' is already used in the destination. Re-run with -AliasPrefix"
        continue
    }
    elseif ($PSCmdlet.ShouldProcess($primary, "Create $(if ($isRoom) { 'room' } else { 'equipment' }) mailbox")) {
        try {
            $name = $r.DisplayName
            if (Get-Recipient -Identity $name -ErrorAction SilentlyContinue) { $name = "$name ($alias)" }
            $nm = @{ Name = $name; Alias = $alias; DisplayName = $r.DisplayName; PrimarySmtpAddress = $primary; ErrorAction = 'Stop' }
            if ($isRoom) { $nm.Room = $true } else { $nm.Equipment = $true }
            # Set the sign-in name, or Exchange puts it on the tenant's default domain (see New-TargetMailboxes).
            $nm.MicrosoftOnlineServicesID = $signIn
            $new = New-Mailbox @nm
            if (-not $new) { throw "New-Mailbox returned nothing for $primary." }
            $id = $primary
            $map[$src] = $id
            Write-Log 'Mailbox' $src $id 'Created' "Sign-in $signIn"
        }
        catch { Write-Log 'Mailbox' $src $primary 'Problem' $_.Exception.Message; continue }
    }
    else {
        Write-Log 'Mailbox' $src $primary 'WhatIf' "Would create $(if ($isRoom) { 'room' } else { 'equipment' }) mailbox, sign-in $signIn"
        $id = $null
    }

    # ---------------- mailbox settings: tag, capacity, office, address list
    $sm = @{ CustomAttribute15 = (Get-Tag $src) }
    if ("$($r.ResourceCapacity)" -match '^\d+$') { $sm.ResourceCapacity = [int]$r.ResourceCapacity }
    if ($r.Office) { $sm.Office = $r.Office }
    $hidden = ConvertTo-Bool $r.HiddenFromAddressListsEnabled
    if ($null -ne $hidden) { $sm.HiddenFromAddressListsEnabled = $hidden }

    # ---------------- calendar processing
    $missing = @()
    $cp = @{}
    if ($r.AutomateProcessing) { $cp.AutomateProcessing = "$($r.AutomateProcessing)" }
    foreach ($n in 'BookingWindowInDays', 'MaximumDurationInMinutes') { if ("$($r.$n)" -match '^\d+$') { $cp[$n] = [int]$r.$n } }
    foreach ($n in 'AllowConflicts', 'AllowRecurringMeetings', 'AllBookInPolicy', 'AllRequestInPolicy', 'AddOrganizerToSubject',
                   'DeleteSubject', 'DeleteComments', 'RemovePrivateProperty', 'ProcessExternalMeetingMessages') {
        $b = ConvertTo-Bool $r.$n
        if ($null -ne $b) { $cp[$n] = $b }
    }
    $delegates = @(Get-MappedList $r.ResourceDelegates ([ref]$missing))
    $bookIn    = @(Get-MappedList $r.BookInPolicy ([ref]$missing))
    $requestIn = @(Get-MappedList $r.RequestInPolicy ([ref]$missing))
    if ($delegates) { $cp.ResourceDelegates = $delegates }
    if ($bookIn)    { $cp.BookInPolicy = $bookIn }
    if ($requestIn) { $cp.RequestInPolicy = $requestIn }
    $detail = "Auto: $($cp.AutomateProcessing); window $($cp.BookingWindowInDays) days; conflicts $($cp.AllowConflicts); delegates $($delegates.Count); book-in $($bookIn.Count)"
    if ($missing) { $detail += "; not mapped, left out: $(($missing | Select-Object -Unique) -join ', ')" }
    if ($r.BookInPolicy -and -not $bookIn -and $cp.AllBookInPolicy -eq $false) { $detail += '; nobody left in BookInPolicy, so requests go to delegates' }

    if (-not $id) {
        Write-Log 'Calendar' $src $primary 'WhatIf' "Would set capacity $($sm.ResourceCapacity) and $detail"
        continue
    }
    if ($PSCmdlet.ShouldProcess($id, 'Set capacity, tag and calendar processing')) {
        try {
            Set-Mailbox -Identity $id @sm -ErrorAction Stop
            Set-CalendarProcessing -Identity $id @cp -ErrorAction Stop
            Write-Log 'Calendar' $src $id 'Set' $detail
        }
        catch {
            # A brand-new mailbox can take a few minutes before Set-* commands find it.
            if ($_.Exception.Message -match "couldn't be found|could not be found|not found") { Write-Log 'Calendar' $src $id 'Pending' 'Mailbox not ready yet. Re-run in 15 minutes' }
            else { Write-Log 'Calendar' $src $id 'Problem' $_.Exception.Message }
        }
    }
    else { Write-Log 'Calendar' $src $id 'WhatIf' "Would set capacity $($sm.ResourceCapacity) and $detail" }
}

# ------------------------------------------------------------ report
$file = ".\TargetResourceMailboxes-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
$log | Export-Csv $file -NoTypeInformation -WhatIf:$false
Write-Host ''
$log | Group-Object Status | Sort-Object Name | ForEach-Object { Write-Host ("{0,-8} {1}" -f $_.Name, $_.Count) }
Write-Host "Report: $file" -ForegroundColor Cyan
if ($log | Where-Object Status -in 'Problem', 'Pending') { Write-Host 'Some items need a re-run. Anything already done is skipped.' -ForegroundColor Yellow }
if ($log | Where-Object { $_.Detail -like '*not mapped*' }) { Write-Host 'Some delegates or booking policy entries had no mapping. See the Detail column, fix the mapping, and re-run.' -ForegroundColor Yellow }
Write-Host 'Next: re-run 02b New-UserMapping so the rooms are in UserMapping.csv for 07 and 09.' -ForegroundColor Cyan
