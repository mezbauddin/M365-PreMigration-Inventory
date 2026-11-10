<#
.SYNOPSIS
    Read-only inventory of a Microsoft 365 source tenant before a tenant-to-tenant migration.

.DESCRIPTION
    TL;DR: Run once as a Global Reader (or Global Admin) in the SOURCE tenant. It writes CSV files covering
    users and licences, mailboxes with sizes and archives, mailbox permissions, groups, distribution lists,
    mail contacts, accepted domains, SharePoint sites and OneDrive storage, then prints a summary.
    Since 1.1.0 it also lists public folders, inactive mailboxes, room and equipment mailboxes with their booking
    settings, and the ExchangeGuid, ArchiveGuid and LegacyExchangeDN that Microsoft's native cross-tenant
    mailbox migration needs on each target mail user. Nothing in the tenant is changed.

    Output folder: .\Inventory-<tenant>-<yyyyMMdd-HHmm>\

.PARAMETER SharePointAdminUrl
    The SharePoint admin URL of the SOURCE tenant, with or without https://, for example
    fabrikam-admin.sharepoint.com

.PARAMETER SkipMailboxPermissions
    Skip the Full Access and Send As scan (the slowest part on large tenants).

.EXAMPLE
    .\Invoke-SourceTenantInventory.ps1 -SharePointAdminUrl https://fabrikam-admin.sharepoint.com

.EXAMPLE
    .\Invoke-SourceTenantInventory.ps1 -SharePointAdminUrl fabrikam-admin.sharepoint.com -SkipMailboxPermissions

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.1.0
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: ExchangeOnlineManagement 3.x, Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement,
              Microsoft.Online.SharePoint.PowerShell (Windows PowerShell 5.1 for the SPO module)
    Changelog:
      1.1.0  Adds 12-PublicFolders.csv, 13-MailPublicFolders.csv, 14-InactiveMailboxes.csv and
             15-ResourceMailboxes.csv (rooms and equipment with capacity and calendar processing). 04-Mailboxes.csv
             gains ExchangeGuid, ArchiveGuid, LegacyExchangeDN and X500Addresses columns at the end. Summary gains
             public folder and inactive mailbox counts. Existing files and columns are unchanged.
      1.0.4  Aliases lists secondary smtp: addresses only (case-sensitive match), not the primary address.
      1.0.3  Fix null ExchangeGuid (not in Get-EXOMailbox's default property set); identity fallback to
             object ID or UPN; per-mailbox errors warn instead of stopping the run.
      1.0.2  Module preflight; SPO module loads via -UseWindowsPowerShell on PS 7; SharePoint section is skipped
             with a warning (instead of failing the run) if the SPO module is missing or cannot connect.
      1.0.1  SharePointAdminUrl accepts the bare host (no https://) and a trailing slash.
      1.0.0  Initial release.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^(https://)?[a-z0-9-]+-admin\.sharepoint\.com/?$')]
    [string]$SharePointAdminUrl,

    [switch]$SkipMailboxPermissions
)

$ErrorActionPreference = 'Stop'

# Accept "fabrikam-admin.sharepoint.com" or "https://fabrikam-admin.sharepoint.com/" and normalise both.
$SharePointAdminUrl = 'https://' + ($SharePointAdminUrl.Trim().ToLower() -replace '^https://', '' -replace '/$', '')

function Convert-SizeToBytes {
    # Exchange returns sizes such as "1.234 GB (1,325,000,000 bytes)". Return the byte count as [long].
    param([object]$Size)
    if ($null -eq $Size) { return [long]0 }
    $text = $Size.ToString()
    if ($text -match '\(([\d,]+) bytes\)') { return [long]($Matches[1] -replace ',', '') }
    return [long]0
}

function Write-Step { param([string]$Message) Write-Host ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $Message) -ForegroundColor Cyan }

# ---------------------------------------------------------------- preflight
Write-Step 'Checking modules'
$required = 'ExchangeOnlineManagement', 'Microsoft.Graph.Users', 'Microsoft.Graph.Identity.DirectoryManagement'
$missing = $required | Where-Object { -not (Get-Module -ListAvailable -Name $_) }
if ($missing) {
    throw ("Missing module(s): {0}. Install with: Install-Module {0} -Scope CurrentUser" -f ($missing -join ', '))
}

# SharePoint Online module is optional: without it the script still inventories Entra and Exchange.
$spoAvailable = $false
if (Get-Module -ListAvailable -Name 'Microsoft.Online.SharePoint.PowerShell') {
    try {
        if ($PSVersionTable.PSEdition -eq 'Core') {
            # The SPO module is built for Windows PowerShell 5.1; load it through the compatibility layer in PS 7.
            Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -DisableNameChecking -WarningAction SilentlyContinue
        }
        else {
            Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -WarningAction SilentlyContinue
        }
        $spoAvailable = $true
    }
    catch {
        Write-Warning "SharePoint Online module found but could not load: $($_.Exception.Message)"
    }
}
if (-not $spoAvailable) {
    Write-Warning 'SharePoint Online module not available. SharePoint and OneDrive storage will be skipped. Install with: Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser (Windows only), then re-run.'
}

# ---------------------------------------------------------------- connect
Write-Step 'Connecting to Microsoft Graph and Exchange Online'
Connect-MgGraph -Scopes 'User.Read.All', 'Directory.Read.All', 'Organization.Read.All', 'AuditLog.Read.All' -NoWelcome
Connect-ExchangeOnline -ShowBanner:$false
if ($spoAvailable) {
    Write-Step "Connecting to SharePoint Online ($SharePointAdminUrl)"
    try {
        Connect-SPOService -Url $SharePointAdminUrl
    }
    catch {
        Write-Warning "Could not connect to SharePoint Online: $($_.Exception.Message). SharePoint and OneDrive storage will be skipped."
        $spoAvailable = $false
    }
}

$tenantName = ($SharePointAdminUrl -replace '^https://', '' -replace '-admin\.sharepoint\.com/?$', '')
$outDir = Join-Path -Path (Get-Location) -ChildPath ("Inventory-{0}-{1:yyyyMMdd-HHmm}" -f $tenantName, (Get-Date))
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

# ---------------------------------------------------------------- domains and licences
Write-Step 'Accepted domains'
Get-AcceptedDomain |
    Select-Object DomainName, DomainType, Default |
    Export-Csv (Join-Path $outDir '01-AcceptedDomains.csv') -NoTypeInformation

Write-Step 'Licence SKUs'
$skus = Get-MgSubscribedSku -All
$skuMap = @{}
foreach ($s in $skus) { $skuMap[$s.SkuId.ToString()] = $s.SkuPartNumber }
$skus |
    Select-Object SkuPartNumber, @{n = 'Purchased'; e = { $_.PrepaidUnits.Enabled } }, ConsumedUnits |
    Export-Csv (Join-Path $outDir '02-LicenceSkus.csv') -NoTypeInformation

# ---------------------------------------------------------------- users
Write-Step 'Users'
try {
    $users = Get-MgUser -All -Property Id, DisplayName, UserPrincipalName, Mail, AccountEnabled, UserType, AssignedLicenses, CreatedDateTime, SignInActivity
}
catch {
    # SignInActivity needs an Entra ID P1 licence and AuditLog.Read.All; retry without it.
    Write-Warning 'Sign-in activity not available; continuing without LastSignIn.'
    $users = Get-MgUser -All -Property Id, DisplayName, UserPrincipalName, Mail, AccountEnabled, UserType, AssignedLicenses, CreatedDateTime
}
$users |
    Select-Object DisplayName, UserPrincipalName, Mail, AccountEnabled, UserType, CreatedDateTime,
        @{n = 'Licences'; e = { ($_.AssignedLicenses | ForEach-Object { $skuMap[$_.SkuId.ToString()] }) -join ';' } },
        @{n = 'LastSignIn'; e = { $_.SignInActivity.LastSignInDateTime } } |
    Sort-Object DisplayName |
    Export-Csv (Join-Path $outDir '03-Users.csv') -NoTypeInformation

# ---------------------------------------------------------------- mailboxes
Write-Step 'Mailboxes and sizes (this can take a few minutes)'
# ExchangeGuid is not in Get-EXOMailbox's default (Minimum) property set, so request it explicitly.
# LegacyExchangeDN, ExchangeGuid and ArchiveGuid are what native cross-tenant moves stamp on the target mail user.
$mailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties ExchangeGuid, ArchiveStatus, ArchiveGuid, LegacyExchangeDN, LitigationHoldEnabled, InPlaceHolds, ForwardingSmtpAddress, ForwardingAddress, DeliverToMailboxAndForward, EmailAddresses
Write-Host ("    {0} mailboxes found" -f @($mailboxes).Count)

function Get-MailboxKey {
    # Most reliable identity first; fall back so one odd object never stops the run.
    param($Mailbox)
    if ($Mailbox.ExchangeGuid -and $Mailbox.ExchangeGuid.ToString() -ne '00000000-0000-0000-0000-000000000000') { return $Mailbox.ExchangeGuid.ToString() }
    if ($Mailbox.ExternalDirectoryObjectId) { return [string]$Mailbox.ExternalDirectoryObjectId }
    return [string]$Mailbox.UserPrincipalName
}

$mbxRows = foreach ($m in $mailboxes) {
    $key = Get-MailboxKey $m
    $stats = $null
    try {
        $stats = Get-EXOMailboxStatistics -Identity $key -Properties LastLogonTime -ErrorAction Stop
    }
    catch {
        Write-Warning "No statistics for $($m.PrimarySmtpAddress): $($_.Exception.Message)"
    }
    $archiveBytes = [long]0
    $archiveItems = 0
    if ($m.ArchiveStatus -eq 'Active') {
        $a = Get-EXOMailboxStatistics -Identity $key -Archive -ErrorAction SilentlyContinue
        if ($a) { $archiveBytes = Convert-SizeToBytes $a.TotalItemSize; $archiveItems = $a.ItemCount }
    }
    $primaryBytes = if ($stats) { Convert-SizeToBytes $stats.TotalItemSize } else { [long]0 }
    [pscustomobject]@{
        DisplayName           = $m.DisplayName
        PrimarySmtpAddress    = $m.PrimarySmtpAddress
        UserPrincipalName     = $m.UserPrincipalName
        RecipientTypeDetails  = $m.RecipientTypeDetails
        ItemCount             = if ($stats) { $stats.ItemCount } else { $null }
        SizeGB                = [math]::Round($primaryBytes / 1e9, 2)
        ArchiveStatus         = $m.ArchiveStatus
        ArchiveItemCount      = $archiveItems
        ArchiveSizeGB         = [math]::Round($archiveBytes / 1e9, 2)
        LitigationHold        = $m.LitigationHoldEnabled
        InPlaceHolds          = ($m.InPlaceHolds -join ';')
        ForwardingSmtpAddress = $m.ForwardingSmtpAddress
        ForwardingAddress     = $m.ForwardingAddress
        LastLogonTime         = if ($stats) { $stats.LastLogonTime } else { $null }
        Aliases               = (($m.EmailAddresses | Where-Object { $_ -clike 'smtp:*' }) -join ';')
        # New in 1.1.0, kept at the end so anything reading the older columns is unaffected.
        ExchangeGuid          = "$($m.ExchangeGuid)"
        ArchiveGuid           = $(if ("$($m.ArchiveGuid)" -and "$($m.ArchiveGuid)" -ne '00000000-0000-0000-0000-000000000000') { "$($m.ArchiveGuid)" } else { '' })
        LegacyExchangeDN      = $m.LegacyExchangeDN
        X500Addresses         = (($m.EmailAddresses | Where-Object { "$_" -like 'x500:*' }) -join ';')
    }
}
$mbxRows | Sort-Object SizeGB -Descending |
    Export-Csv (Join-Path $outDir '04-Mailboxes.csv') -NoTypeInformation

# ---------------------------------------------------------------- mailbox permissions
if (-not $SkipMailboxPermissions) {
    Write-Step 'Full Access and Send As permissions'
    $permRows = foreach ($m in $mailboxes) {
        $key = Get-MailboxKey $m
        try {
            Get-EXOMailboxPermission -Identity $key -ErrorAction Stop |
                Where-Object { -not $_.IsInherited -and $_.User -notlike 'NT AUTHORITY\*' -and $_.User -notlike 'S-1-5-*' } |
                ForEach-Object {
                    [pscustomobject]@{ Mailbox = $m.PrimarySmtpAddress; Trustee = $_.User; Right = ($_.AccessRights -join ';') }
                }
            Get-EXORecipientPermission -Identity $key -ErrorAction Stop |
                Where-Object { $_.Trustee -notlike 'NT AUTHORITY\*' } |
                ForEach-Object {
                    [pscustomobject]@{ Mailbox = $m.PrimarySmtpAddress; Trustee = $_.Trustee; Right = 'SendAs' }
                }
        }
        catch {
            Write-Warning "Could not read permissions for $($m.PrimarySmtpAddress): $($_.Exception.Message)"
        }
    }
    $permRows | Export-Csv (Join-Path $outDir '05-MailboxPermissions.csv') -NoTypeInformation
}

# ---------------------------------------------------------------- groups, lists, contacts
Write-Step 'Microsoft 365 groups'
Get-UnifiedGroup -ResultSize Unlimited |
    Select-Object DisplayName, PrimarySmtpAddress, AccessType, SharePointSiteUrl, ResourceProvisioningOptions, WhenCreated |
    Export-Csv (Join-Path $outDir '06-M365Groups.csv') -NoTypeInformation

Write-Step 'Distribution lists and members'
$dlRows = foreach ($dl in (Get-DistributionGroup -ResultSize Unlimited)) {
    $members = Get-DistributionGroupMember -Identity $dl.Identity -ResultSize Unlimited
    [pscustomobject]@{
        DisplayName        = $dl.DisplayName
        PrimarySmtpAddress = $dl.PrimarySmtpAddress
        Type               = $dl.RecipientTypeDetails
        MemberCount        = @($members).Count
        Members            = (($members | ForEach-Object { $_.PrimarySmtpAddress }) -join ';')
        ExternalSenders    = -not $dl.RequireSenderAuthenticationEnabled
    }
}
$dlRows | Export-Csv (Join-Path $outDir '07-DistributionLists.csv') -NoTypeInformation

Write-Step 'Mail contacts'
$contacts = Get-MailContact -ResultSize Unlimited
$contacts |
    Select-Object DisplayName, ExternalEmailAddress, HiddenFromAddressListsEnabled |
    Export-Csv (Join-Path $outDir '08-MailContacts.csv') -NoTypeInformation

Write-Step 'Mail flow rules and EWS setting'
Get-TransportRule | Select-Object Name, State, Priority, Description |
    Export-Csv (Join-Path $outDir '09-TransportRules.csv') -NoTypeInformation
Get-OrganizationConfig | Select-Object EWSEnabled, @{n = 'EWSAllowList'; e = { $_.EWSAllowList -join ';' } } |
    Export-Csv (Join-Path $outDir '10-EwsConfig.csv') -NoTypeInformation
$publicFolderMailboxes = @(Get-Mailbox -PublicFolder -ResultSize Unlimited -ErrorAction SilentlyContinue)

# ---------------------------------------------------------------- public folders
# No copy tool moves public folders by default, and native cross-tenant migration doesn't move them at all.
$publicFolders = @()
$mailPublicFolders = @()
if ($publicFolderMailboxes.Count) {
    Write-Step ("Public folders ({0} public folder mailbox(es))" -f $publicFolderMailboxes.Count)
    try {
        $publicFolders = @(Get-PublicFolder -Identity '' -Recurse -ResultSize Unlimited -ErrorAction Stop |
            Where-Object { $_.ParentPath })   # the root (IPM_SUBTREE) has no parent path
    }
    catch { Write-Warning "Could not read public folders: $($_.Exception.Message)" }
    $publicFolders |
        Select-Object Name, ParentPath, FolderClass, MailEnabled, ContentMailboxName, HasSubfolders |
        Export-Csv (Join-Path $outDir '12-PublicFolders.csv') -NoTypeInformation
    try { $mailPublicFolders = @(Get-MailPublicFolder -ResultSize Unlimited -ErrorAction Stop) }
    catch { Write-Warning "Could not read mail-enabled public folders: $($_.Exception.Message)" }
    $mailPublicFolders |
        Select-Object DisplayName, PrimarySmtpAddress, Alias, HiddenFromAddressListsEnabled,
            @{n = 'Aliases'; e = { ($_.EmailAddresses | Where-Object { $_ -clike 'smtp:*' }) -join ';' } } |
        Export-Csv (Join-Path $outDir '13-MailPublicFolders.csv') -NoTypeInformation
}
else {
    Write-Step 'Public folders: none (no public folder mailboxes)'
}

# ---------------------------------------------------------------- inactive mailboxes
# Mailboxes of deleted users kept by a hold. They don't show in Get-EXOMailbox, they don't move, and they matter at decommission.
Write-Step 'Inactive mailboxes'
$inactive = @()
try { $inactive = @(Get-Mailbox -InactiveMailboxOnly -ResultSize Unlimited -ErrorAction Stop) }
catch { Write-Warning "Could not read inactive mailboxes: $($_.Exception.Message)" }
$inactive |
    Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails, WhenSoftDeleted, LitigationHoldEnabled,
        @{n = 'InPlaceHolds'; e = { $_.InPlaceHolds -join ';' } }, ArchiveStatus,
        @{n = 'ExchangeGuid'; e = { "$($_.ExchangeGuid)" } } |
    Export-Csv (Join-Path $outDir '14-InactiveMailboxes.csv') -NoTypeInformation

# ---------------------------------------------------------------- room and equipment mailboxes
# 02-TargetMailboxes\New-TargetResourceMailboxes.ps1 reads this file to rebuild them with the same booking rules.
Write-Step 'Room and equipment mailboxes, with calendar processing'
$recipientCache = @{}
function Resolve-RecipientSmtp {
    # Calendar processing returns delegates and policies as names or legacy DNs. Turn each into an SMTP address.
    param([object]$Id)
    if (-not $Id) { return $null }
    $k = "$Id"
    if ($recipientCache.ContainsKey($k)) { return $recipientCache[$k] }
    $r = Get-EXORecipient -Identity $k -ErrorAction SilentlyContinue
    $recipientCache[$k] = if ($r) { "$($r.PrimarySmtpAddress)".ToLower() } else { "unresolved:$k" }
    return $recipientCache[$k]
}
$resourceRows = foreach ($m in @($mailboxes | Where-Object { $_.RecipientTypeDetails -in 'RoomMailbox', 'EquipmentMailbox' })) {
    $key = Get-MailboxKey $m
    try {
        $full = Get-Mailbox -Identity $key -ErrorAction Stop
        $cp = Get-CalendarProcessing -Identity $key -ErrorAction Stop
    }
    catch {
        Write-Warning "Could not read booking settings for $($m.PrimarySmtpAddress): $($_.Exception.Message)"
        continue
    }
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
        BookInPolicy                   = (@($cp.BookInPolicy | ForEach-Object { Resolve-RecipientSmtp $_ }) -join ';')
        RequestInPolicy                = (@($cp.RequestInPolicy | ForEach-Object { Resolve-RecipientSmtp $_ }) -join ';')
        ResourceDelegates              = (@($cp.ResourceDelegates | ForEach-Object { Resolve-RecipientSmtp $_ }) -join ';')
        AddOrganizerToSubject          = $cp.AddOrganizerToSubject
        DeleteSubject                  = $cp.DeleteSubject
        DeleteComments                 = $cp.DeleteComments
        RemovePrivateProperty          = $cp.RemovePrivateProperty
        ProcessExternalMeetingMessages = $cp.ProcessExternalMeetingMessages
    }
}
@($resourceRows) | Export-Csv (Join-Path $outDir '15-ResourceMailboxes.csv') -NoTypeInformation

# ---------------------------------------------------------------- SharePoint and OneDrive
$siteRows = @()
if ($spoAvailable) {
    Write-Step 'SharePoint sites and OneDrive accounts (storage includes versions and recycle bin)'
    $sites = Get-SPOSite -Limit All -IncludePersonalSite $true
    $siteRows = $sites | Select-Object Url, Title, Template, Owner,
        @{n = 'StorageGB'; e = { [math]::Round($_.StorageUsageCurrent / 1024, 2) } },
        LastContentModifiedDate, LockState, SharingCapability,
        @{n = 'IsOneDrive'; e = { $_.Url -like '*-my.sharepoint.com/personal/*' } }
    $siteRows | Sort-Object StorageGB -Descending |
        Export-Csv (Join-Path $outDir '11-SitesAndOneDrive.csv') -NoTypeInformation
}
else {
    Write-Step 'Skipping SharePoint and OneDrive (module not loaded)'
}

# ---------------------------------------------------------------- summary
$summary = [ordered]@{
    Tenant                 = $tenantName
    RunAt                  = (Get-Date).ToString('u')
    EnabledMemberUsers     = @($users | Where-Object { $_.AccountEnabled -and $_.UserType -eq 'Member' }).Count
    UserMailboxes          = @($mbxRows | Where-Object RecipientTypeDetails -eq 'UserMailbox').Count
    SharedMailboxes        = @($mbxRows | Where-Object RecipientTypeDetails -eq 'SharedMailbox').Count
    RoomEquipmentMailboxes = @($mbxRows | Where-Object { $_.RecipientTypeDetails -in 'RoomMailbox', 'EquipmentMailbox' }).Count
    MailboxTotalGB         = [math]::Round(($mbxRows | Measure-Object SizeGB -Sum).Sum, 2)
    ArchiveTotalGB         = [math]::Round(($mbxRows | Measure-Object ArchiveSizeGB -Sum).Sum, 2)
    LargestMailboxGB       = ($mbxRows | Measure-Object SizeGB -Maximum).Maximum
    MailboxesOnHold        = @($mbxRows | Where-Object { $_.LitigationHold -or $_.InPlaceHolds }).Count
    PublicFolderMailboxes  = $publicFolderMailboxes.Count
    DistributionLists      = @($dlRows).Count
    MailContacts           = @($contacts).Count
    SharePointSites        = @($siteRows | Where-Object { -not $_.IsOneDrive }).Count
    SharePointTotalGB      = [math]::Round(($siteRows | Where-Object { -not $_.IsOneDrive } | Measure-Object StorageGB -Sum).Sum, 2)
    OneDriveAccounts       = @($siteRows | Where-Object IsOneDrive).Count
    OneDriveTotalGB        = [math]::Round(($siteRows | Where-Object IsOneDrive | Measure-Object StorageGB -Sum).Sum, 2)
    # New in 1.1.0, appended so the existing columns keep their order.
    PublicFolders          = $publicFolders.Count
    MailPublicFolders      = $mailPublicFolders.Count
    InactiveMailboxes      = $inactive.Count
}
[pscustomobject]$summary | Export-Csv (Join-Path $outDir '00-Summary.csv') -NoTypeInformation
[pscustomobject]$summary | Format-List

Write-Step "Done. Files written to $outDir"
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-MgGraph | Out-Null
if ($spoAvailable) { Disconnect-SPOService -ErrorAction SilentlyContinue }
