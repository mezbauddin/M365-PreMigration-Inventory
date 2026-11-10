<#
.SYNOPSIS
    Exports the mail configuration a migration tool doesn't copy from the source tenant, and rebuilds it in the
    destination: mailbox permissions, calendar and folder permissions, forwarding, distribution lists, mail contacts.
    Transport rules are exported for review; inbox rules and automatic replies are listed to check after cutover.

.DESCRIPTION
    TL;DR: Mailbox copy tools move the mail. They don't move who can open whose mailbox, who can send as whom,
    who can see or edit whose calendar, forwarding, distribution lists or contacts. Run -Export against the source, then -Import against the destination
    with a mapping CSV (Source,Destination). Import is safe to re-run: anything it already created is recognised and
    left alone. Everything supports -WhatIf.

    -Export writes a folder of CSVs:
      Permissions.csv            Full Access, Send As and Send on Behalf, per mailbox
      FolderPermissions.csv      Calendar permissions (and Inbox, Contacts with -IncludeFolders), including calendar
                                 delegates. Default and Anonymous only when changed from the defaults
      CheckAfterCutover.csv      Report only: inbox rule counts (and rules that forward) and automatic reply state,
                                 per mailbox that has either. Not imported
      Forwarding.csv             Mailbox forwarding, marked internal or external
      DistributionGroups.csv     Distribution lists and mail-enabled security groups, with settings and owners
      DistributionGroupMembers.csv
      DynamicGroups.csv          Dynamic distribution lists and their filters (rebuild by hand)
      MailContacts.csv
      TransportRules.csv         Rule summary. TransportRules.xml holds the full export
    -Import rebuilds, in this order: contacts, distribution lists, their members, permissions, folder permissions,
    forwarding.

.PARAMETER Export
    Read the source tenant and write the export folder. Changes nothing.

.PARAMETER Import
    Rebuild the configuration in the destination tenant from an export folder.

.PARAMETER TenantId
    Optional tenant ID (GUID). The script stops if Exchange Online is signed in to any other tenant.

.PARAMETER AdminUpn
    Optional admin sign-in name, e.g. admin@contoso.com. If Exchange Online is already signed in as someone else, or
    to another tenant, the script disconnects and signs in again as this account. Use it whenever you have both
    tenants open on one PC.

.PARAMETER OutputFolder
    With -Export: where to write. Defaults to .\MailConfig-<domain>-<date>.

.PARAMETER IncludeFolders
    With -Export: also export permissions on these default folders, as well as Calendar. Inbox, Contacts.

.PARAMETER SkipFolderPermissions
    With -Export: don't read folder permissions. Faster on big tenants; FolderPermissions.csv isn't written.

.PARAMETER SkipMailboxChecks
    With -Export: don't read inbox rules and automatic replies. CheckAfterCutover.csv isn't written.

.PARAMETER ExportFolder
    With -Import: the folder -Export wrote.

.PARAMETER MappingFile
    With -Import: CSV with Source and Destination columns, mapping each source mailbox address to its destination
    identity (UPN or address). The same file New-TargetMailboxes users can build for SharePoint and Teams mapping.

.PARAMETER AddressDomain
    With -Import: accepted domain in the destination for the primary address of new distribution lists, e.g. a
    staging domain. Source domains are added at cutover.

.PARAMETER AliasPrefix
    With -Import: prefix for new distribution list aliases, for when the destination already uses names like 'info'.

.PARAMETER Include
    With -Import: which parts to rebuild. Default: Contacts, DistributionGroups, Permissions, FolderPermissions,
    Forwarding.

.PARAMETER IncludeExternalForwarding
    With -Import: also recreate forwarding to addresses outside the organisation. Off by default, because it sends
    mail out of the tenant and outbound spam policy may block it anyway.

.EXAMPLE
    .\Copy-MailConfig.ps1 -Export -TenantId <source-guid>

.EXAMPLE
    .\Copy-MailConfig.ps1 -Import -TenantId <dest-guid> -ExportFolder .\MailConfig-fabrikam.com-20261109-1100 `
        -MappingFile .\UserMapping.csv -AddressDomain contoso.com -WhatIf
.EXAMPLE
    .\Copy-MailConfig.ps1 -Export -TenantId <source-guid> -AdminUpn admin@fabrikam.com

.EXAMPLE
    .\Copy-MailConfig.ps1 -Export -TenantId <source-guid> -AdminUpn admin@fabrikam.com -IncludeFolders Inbox, Contacts

.EXAMPLE
    .\Copy-MailConfig.ps1 -Import -TenantId <dest-guid> -AdminUpn admin@contoso.onmicrosoft.com -ExportFolder .\MailConfig-fabrikam.com-20261109-1100 `
        -MappingFile .\UserMapping.csv -AddressDomain contoso.com -Include FolderPermissions -WhatIf

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.2.0
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: ExchangeOnlineManagement 3.x. Exchange Administrator. Objects this script creates are tagged with
              CustomAttribute15 = 'T2T:<source address>' so re-runs and the cutover step can find them.
    Folder permissions are rebuilt with Add-MailboxFolderPermission or Set-MailboxFolderPermission on the
    destination's own folder name (Calendar can be called Kalender or Agenda), found with
    Get-EXOMailboxFolderStatistics -FolderScope.
    Inbox rules and automatic replies are content of the mailbox: native cross-tenant moves carry them, ShareGate
    copies inbox rules (except a few rule options) and some automatic reply settings, and many other tools copy
    neither. CheckAfterCutover.csv is the list to check against, whatever moved the mail.
    Transport rules are exported, never imported. Import-TransportRuleCollection is retired in Exchange Online (it
    used to replace the whole rule collection), and rules usually point at source-only recipients and domains.
    TransportRules.xml holds a New-TransportRule command per rule: review each one, fix the addresses, and run it
    with -Enabled $false until you've tested it.
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Export')]
param(
    [Parameter(ParameterSetName = 'Export', Mandatory)][switch]$Export,
    [Parameter(ParameterSetName = 'Import', Mandatory)][switch]$Import,
    [ValidatePattern('^[0-9a-fA-F-]{36}$')][string]$TenantId,
    [string]$AdminUpn,

    [Parameter(ParameterSetName = 'Export')][string]$OutputFolder,
    [Parameter(ParameterSetName = 'Export')][ValidateSet('Inbox', 'Contacts')][string[]]$IncludeFolders = @(),
    [Parameter(ParameterSetName = 'Export')][switch]$SkipFolderPermissions,
    [Parameter(ParameterSetName = 'Export')][switch]$SkipMailboxChecks,

    [Parameter(ParameterSetName = 'Import', Mandatory)][string]$ExportFolder,
    [Parameter(ParameterSetName = 'Import', Mandatory)][string]$MappingFile,
    [Parameter(ParameterSetName = 'Import', Mandatory)][string]$AddressDomain,
    [Parameter(ParameterSetName = 'Import')][string]$AliasPrefix = '',
    [Parameter(ParameterSetName = 'Import')][ValidateSet('Contacts', 'DistributionGroups', 'Permissions', 'FolderPermissions', 'Forwarding')]
    [string[]]$Include = @('Contacts', 'DistributionGroups', 'Permissions', 'FolderPermissions', 'Forwarding'),
    [Parameter(ParameterSetName = 'Import')][switch]$IncludeExternalForwarding
)

$ErrorActionPreference = 'Stop'

# ------------------------------------------------------------ connection and tenant guard
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
if ($TenantId -and $conn.TenantID -ne $TenantId) {
    throw "Exchange Online is signed in to tenant $($conn.TenantID), not $TenantId. Run Disconnect-ExchangeOnline and sign in to the right tenant, or pass -AdminUpn."
}
if ($AdminUpn -and "$($conn.UserPrincipalName)" -ne $AdminUpn) {
    throw "Exchange Online signed in as $($conn.UserPrincipalName), not $AdminUpn. Pick the right account in the sign-in box, or use a private browser window."
}
Write-Host "Signed in as $($conn.UserPrincipalName) to tenant $($conn.TenantID)" -ForegroundColor Cyan
$accepted = @((Get-AcceptedDomain).DomainName | ForEach-Object { $_.ToLower() })
$tenantName = (Get-AcceptedDomain | Where-Object Default).DomainName
Write-Host "Connected to $tenantName" -ForegroundColor Cyan

function Get-SmtpOnly($a) { if ($a) { ("$a" -replace '^smtp:', '').Trim().ToLower() } }
function Get-Tag($src) { "T2T:$($src.ToLower())" }

# Folder names are localised (Calendar, Kalender, Agenda), so find the real path from the folder type.
$folderPathCache = @{}
function Get-DefaultFolderPath([string]$Mailbox, [string]$FolderType) {
    $k = "$Mailbox|$FolderType"
    if (-not $folderPathCache.ContainsKey($k)) {
        $f = Get-EXOMailboxFolderStatistics -Identity $Mailbox -FolderScope $FolderType -ErrorAction SilentlyContinue |
            Where-Object { "$($_.FolderType)" -eq $FolderType } | Select-Object -First 1
        $folderPathCache[$k] = if ($f) { "$($f.FolderPath)" -replace '/', '\' } else { $null }
    }
    $folderPathCache[$k]
}
# What Default and Anonymous have on a fresh mailbox. Anything else was changed on purpose, so it's exported.
$defaultRights = @{ 'Calendar|Default' = 'AvailabilityOnly'; 'Calendar|Anonymous' = 'None'; 'Inbox|Default' = 'None'; 'Inbox|Anonymous' = 'None'; 'Contacts|Default' = 'None'; 'Contacts|Anonymous' = 'None' }

# ============================================================ EXPORT
if ($Export) {
    if (-not $OutputFolder) { $OutputFolder = ".\MailConfig-$tenantName-$(Get-Date -Format 'yyyyMMdd-HHmm')" }
    New-Item -ItemType Directory -Path $OutputFolder -Force -WhatIf:$false | Out-Null

    $cache = @{}
    function Resolve-Smtp($id) {
        if (-not $id) { return $null }
        $k = "$id"
        if ($cache.ContainsKey($k)) { return $cache[$k] }
        $r = Get-EXORecipient -Identity $k -ErrorAction SilentlyContinue
        $cache[$k] = if ($r) { "$($r.PrimarySmtpAddress)".ToLower() } else { "unresolved:$k" }
        return $cache[$k]
    }

    $mbx = @(Get-EXOMailbox -ResultSize Unlimited -Properties GrantSendOnBehalfTo, ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward |
        Where-Object RecipientTypeDetails -ne 'DiscoveryMailbox')
    Write-Host "$($mbx.Count) mailboxes. Reading permissions, this takes a while on big tenants." -ForegroundColor Cyan

    $perm = [System.Collections.Generic.List[object]]::new()
    $fwd  = [System.Collections.Generic.List[object]]::new()
    $folderPerm = [System.Collections.Generic.List[object]]::new()
    $checks = [System.Collections.Generic.List[object]]::new()
    $folderTypes = @('Calendar') + @($IncludeFolders)
    $i = 0
    foreach ($m in $mbx) {
        $i++; Write-Progress -Activity 'Permissions' -Status "$($m.PrimarySmtpAddress)" -PercentComplete (100 * $i / $mbx.Count)
        $addr = "$($m.PrimarySmtpAddress)".ToLower()
        $id = $m.ExternalDirectoryObjectId

        Get-EXOMailboxPermission -Identity $id -ErrorAction SilentlyContinue |
            Where-Object { -not $_.IsInherited -and -not $_.Deny -and $_.AccessRights -contains 'FullAccess' -and
                           $_.User -notlike 'NT AUTHORITY\*' -and $_.User -notlike 'S-1-5-*' } |
            ForEach-Object { $perm.Add([pscustomobject]@{ Mailbox = $addr; Right = 'FullAccess'; Trustee = (Resolve-Smtp $_.User) }) }

        Get-EXORecipientPermission -Identity $id -ErrorAction SilentlyContinue |
            Where-Object { $_.AccessRights -contains 'SendAs' -and $_.Trustee -notlike 'NT AUTHORITY\*' -and $_.Trustee -notlike 'S-1-5-*' } |
            ForEach-Object { $perm.Add([pscustomobject]@{ Mailbox = $addr; Right = 'SendAs'; Trustee = (Resolve-Smtp $_.Trustee) }) }

        foreach ($d in @($m.GrantSendOnBehalfTo)) {
            if ($d) { $perm.Add([pscustomobject]@{ Mailbox = $addr; Right = 'SendOnBehalf'; Trustee = (Resolve-Smtp $d) }) }
        }

        if ($m.ForwardingAddress) {
            $fwd.Add([pscustomobject]@{ Mailbox = $addr; ForwardTo = (Resolve-Smtp $m.ForwardingAddress); Kind = 'Internal'; KeepCopy = $m.DeliverToMailboxAndForward })
        }
        if ($m.ForwardingSmtpAddress) {
            $to = Get-SmtpOnly $m.ForwardingSmtpAddress
            $kind = if ($accepted -contains $to.Split('@')[-1]) { 'Internal' } else { 'External' }
            $fwd.Add([pscustomobject]@{ Mailbox = $addr; ForwardTo = $to; Kind = $kind; KeepCopy = $m.DeliverToMailboxAndForward })
        }

        # ---- calendar and folder permissions
        if (-not $SkipFolderPermissions) {
            foreach ($ft in $folderTypes) {
                $path = Get-DefaultFolderPath $addr $ft
                if (-not $path) { continue }
                foreach ($fp in @(Get-MailboxFolderPermission -Identity "${addr}:$path" -ErrorAction SilentlyContinue)) {
                    $userText = if ($fp.User.DisplayName) { "$($fp.User.DisplayName)" } else { "$($fp.User)" }
                    $userType = "$($fp.User.UserType)"
                    $rights = (@($fp.AccessRights) | ForEach-Object { "$_" }) -join ';'
                    if ($userType -in 'Default', 'Anonymous' -or $userText -in 'Default', 'Anonymous') {
                        $who = if ($userType -in 'Default', 'Anonymous') { $userType } else { $userText }
                        if ($defaultRights["$ft|$who"] -eq $rights) { continue }   # unchanged from a fresh mailbox
                        $trustee = $who
                    }
                    elseif ($userType -eq 'External') { $trustee = "external:$userText" }
                    else {
                        $smtp = "$($fp.User.RecipientPrincipal.PrimarySmtpAddress)"
                        if (-not $smtp) { $smtp = "$($fp.User.ADRecipient.PrimarySmtpAddress)" }
                        $trustee = if ($smtp) { $smtp.ToLower() } else { Resolve-Smtp $userText }
                    }
                    $flags = "$($fp.SharingPermissionFlags)"
                    $folderPerm.Add([pscustomobject]@{ Mailbox = $addr; Folder = $ft; FolderPath = $path; Trustee = $trustee; AccessRights = $rights
                            SharingPermissionFlags = $(if ($flags -and $flags -ne 'None') { $flags } else { '' }) })
                }
            }
        }

        # ---- report only: inbox rules and automatic replies, to check after cutover
        if (-not $SkipMailboxChecks) {
            $rules = @(Get-InboxRule -Mailbox $addr -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)
            $forwarding = @($rules | Where-Object { $_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo })
            $ar = Get-MailboxAutoReplyConfiguration -Identity $addr -ErrorAction SilentlyContinue
            $arState = if ($ar) { "$($ar.AutoReplyState)" } else { 'Unknown' }
            if ($rules.Count -or ($arState -and $arState -ne 'Disabled')) {
                $what = @()
                if ($rules.Count) { $what += "$($rules.Count) inbox rule(s): check they arrived and still point at the right people and folders" }
                if ($forwarding.Count) { $what += "$($forwarding.Count) forward or redirect: check the address, and outbound spam policy" }
                if ($arState -eq 'Enabled') { $what += 'automatic reply on: check it in the destination' }
                if ($arState -eq 'Scheduled') { $what += "automatic reply scheduled until $($ar.EndTime)" }
                $checks.Add([pscustomobject]@{
                        Mailbox                   = $addr
                        Type                      = $m.RecipientTypeDetails
                        InboxRules                = $rules.Count
                        EnabledRules              = @($rules | Where-Object Enabled).Count
                        RulesThatForwardOrRedirect = $forwarding.Count
                        AutoReplyState            = $arState
                        AutoReplyEnd              = $(if ($arState -eq 'Scheduled') { $ar.EndTime } else { '' })
                        Check                     = ($what -join '; ')
                    })
            }
        }
    }
    Write-Progress -Activity 'Permissions' -Completed
    $perm | Export-Csv (Join-Path $OutputFolder 'Permissions.csv') -NoTypeInformation -WhatIf:$false
    $fwd  | Export-Csv (Join-Path $OutputFolder 'Forwarding.csv') -NoTypeInformation -WhatIf:$false
    if (-not $SkipFolderPermissions) {
        # Written even when empty, so -Import can tell "none" from "not exported".
        if ($folderPerm.Count) { $folderPerm | Export-Csv (Join-Path $OutputFolder 'FolderPermissions.csv') -NoTypeInformation -WhatIf:$false }
        else { Set-Content -Path (Join-Path $OutputFolder 'FolderPermissions.csv') -Value '"Mailbox","Folder","FolderPath","Trustee","AccessRights","SharingPermissionFlags"' -WhatIf:$false }
    }
    if (-not $SkipMailboxChecks) {
        $checks | Export-Csv (Join-Path $OutputFolder 'CheckAfterCutover.csv') -NoTypeInformation -WhatIf:$false
    }

    $groups = @(Get-DistributionGroup -ResultSize Unlimited)
    $members = [System.Collections.Generic.List[object]]::new()
    $groupRows = foreach ($g in $groups) {
        $gAddr = "$($g.PrimarySmtpAddress)".ToLower()
        Get-DistributionGroupMember -Identity $g.ExternalDirectoryObjectId -ResultSize Unlimited -ErrorAction SilentlyContinue | ForEach-Object {
            $members.Add([pscustomobject]@{ Group = $gAddr; Member = "$($_.PrimarySmtpAddress)".ToLower(); MemberName = $_.Name; MemberType = $_.RecipientTypeDetails })
        }
        [pscustomobject]@{
            DisplayName                        = $g.DisplayName
            Alias                              = $g.Alias
            PrimarySmtpAddress                 = $gAddr
            Security                           = ("$($g.GroupType)" -match 'SecurityEnabled')
            Owners                             = (@($g.ManagedBy | ForEach-Object { Resolve-Smtp $_ }) -join ';')
            RequireSenderAuthenticationEnabled = $g.RequireSenderAuthenticationEnabled
            HiddenFromAddressListsEnabled      = $g.HiddenFromAddressListsEnabled
            MemberJoinRestriction              = $g.MemberJoinRestriction
            MemberDepartRestriction            = $g.MemberDepartRestriction
            AllSmtpAddresses                   = (@($g.EmailAddresses | Where-Object { $_ -like 'smtp:*' } | ForEach-Object { Get-SmtpOnly $_ } | Select-Object -Unique) -join ';')
        }
    }
    $groupRows | Export-Csv (Join-Path $OutputFolder 'DistributionGroups.csv') -NoTypeInformation -WhatIf:$false
    $members   | Export-Csv (Join-Path $OutputFolder 'DistributionGroupMembers.csv') -NoTypeInformation -WhatIf:$false

    Get-DynamicDistributionGroup -ResultSize Unlimited | ForEach-Object {
        [pscustomobject]@{ DisplayName = $_.DisplayName; Alias = $_.Alias; PrimarySmtpAddress = "$($_.PrimarySmtpAddress)".ToLower(); RecipientFilter = $_.RecipientFilter }
    } | Export-Csv (Join-Path $OutputFolder 'DynamicGroups.csv') -NoTypeInformation -WhatIf:$false

    Get-MailContact -ResultSize Unlimited | ForEach-Object {
        [pscustomobject]@{
            DisplayName                   = $_.DisplayName
            Name                          = $_.Name
            Alias                         = $_.Alias
            ExternalEmailAddress          = Get-SmtpOnly $_.ExternalEmailAddress
            PrimarySmtpAddress            = "$($_.PrimarySmtpAddress)".ToLower()
            HiddenFromAddressListsEnabled = $_.HiddenFromAddressListsEnabled
        }
    } | Export-Csv (Join-Path $OutputFolder 'MailContacts.csv') -NoTypeInformation -WhatIf:$false

    $rules = @(Get-TransportRule -ResultSize Unlimited)
    $rules | Select-Object Priority, Name, State, Mode, Description |
        Export-Csv (Join-Path $OutputFolder 'TransportRules.csv') -NoTypeInformation -WhatIf:$false
    try {
        $collection = Export-TransportRuleCollection
        [IO.File]::WriteAllBytes((Join-Path (Resolve-Path $OutputFolder) 'TransportRules.xml'), $collection.FileData)
    }
    catch { Write-Warning "Couldn't export the full rule collection: $($_.Exception.Message). TransportRules.csv still lists them." }

    $unresolved = @($perm | Where-Object Trustee -like 'unresolved:*').Count
    Write-Host ''
    Write-Host "Export written to $(Resolve-Path $OutputFolder)" -ForegroundColor Green
    Write-Host ("  {0} permissions, {1} forwards ({2} external), {3} distribution lists with {4} members, {5} dynamic lists, {6} contacts, {7} transport rules" -f
        $perm.Count, $fwd.Count, @($fwd | Where-Object Kind -eq 'External').Count, @($groupRows).Count, $members.Count,
        @(Import-Csv (Join-Path $OutputFolder 'DynamicGroups.csv')).Count, @(Import-Csv (Join-Path $OutputFolder 'MailContacts.csv')).Count, $rules.Count)
    if (-not $SkipFolderPermissions) {
        Write-Host ("  {0} folder permissions on {1} ({2} calendar delegates, {3} external)" -f $folderPerm.Count, ($folderTypes -join ', '),
            @($folderPerm | Where-Object { $_.SharingPermissionFlags -match 'Delegate' }).Count, @($folderPerm | Where-Object Trustee -like 'external:*').Count)
    }
    if (-not $SkipMailboxChecks) {
        Write-Host ("  {0} mailbox(es) with inbox rules or an automatic reply to check after cutover. See CheckAfterCutover.csv" -f $checks.Count)
    }
    if ($unresolved) { Write-Host "  $unresolved permission(s) point at deleted or unknown accounts (Trustee 'unresolved:'). Import skips them." -ForegroundColor Yellow }
    Write-Host 'The CSVs hold names and addresses. Keep them out of source control and shared drives.' -ForegroundColor Yellow
    return
}

# ============================================================ IMPORT
if ($accepted -notcontains $AddressDomain.ToLower()) { throw "$AddressDomain isn't an accepted domain in $tenantName." }
foreach ($f in 'Permissions.csv', 'Forwarding.csv', 'DistributionGroups.csv', 'DistributionGroupMembers.csv', 'MailContacts.csv') {
    if (-not (Test-Path (Join-Path $ExportFolder $f))) { throw "$f missing from $ExportFolder. Point -ExportFolder at the folder -Export wrote." }
}

$map = @{}
Import-Csv $MappingFile | ForEach-Object {
    if (-not $_.Source -or -not $_.Destination) { throw "$MappingFile needs Source and Destination columns." }
    $map[$_.Source.Trim().ToLower()] = $_.Destination.Trim()
}
Write-Host "$($map.Count) mailbox mappings loaded." -ForegroundColor Cyan

$log = [System.Collections.Generic.List[object]]::new()
function Write-Log($Area, $Object, $Status, $Detail) {
    $log.Add([pscustomobject]@{ Area = $Area; Object = $Object; Status = $Status; Detail = $Detail })
    $c = @{ Created = 'Green'; Added = 'Green'; Set = 'Green'; Exists = 'Gray'; Skipped = 'Yellow'; Problem = 'Red'; WhatIf = 'Cyan' }[$Status]
    Write-Host ("  [{0}] {1}: {2} {3}" -f $Status, $Area, $Object, $Detail) -ForegroundColor $c
}
function Get-Ours($src) { Get-Recipient -Filter "CustomAttribute15 -eq '$(Get-Tag $src)'" -ErrorAction SilentlyContinue | Select-Object -First 1 }
function Get-Mapped($src) { if ($src) { $map["$src".ToLower()] } }

# ---------------------------------------------------------- contacts
if ($Include -contains 'Contacts') {
    Write-Host 'Contacts' -ForegroundColor Cyan
    foreach ($c in Import-Csv (Join-Path $ExportFolder 'MailContacts.csv')) {
        $src = $c.PrimarySmtpAddress
        $ours = Get-Ours $src
        if ($ours) { $map[$src] = "$($ours.PrimarySmtpAddress)"; Write-Log 'Contact' $c.DisplayName 'Exists' ''; continue }
        $same = Get-Recipient -Identity $c.ExternalEmailAddress -ErrorAction SilentlyContinue
        if ($same) { $map[$src] = "$($same.PrimarySmtpAddress)"; Write-Log 'Contact' $c.DisplayName 'Exists' "Destination already has $($c.ExternalEmailAddress) as a $($same.RecipientTypeDetails). Using it"; continue }
        if ($PSCmdlet.ShouldProcess($c.ExternalEmailAddress, 'Create mail contact')) {
            try {
                $name = $c.Name
                if (Get-Recipient -Identity $name -ErrorAction SilentlyContinue) { $name = "$name (migrated)" }
                $new = New-MailContact -Name $name -DisplayName $c.DisplayName -ExternalEmailAddress $c.ExternalEmailAddress -ErrorAction Stop
                Set-MailContact -Identity $new.Identity -CustomAttribute15 (Get-Tag $src) -HiddenFromAddressListsEnabled ([bool]::Parse("$($c.HiddenFromAddressListsEnabled)")) -ErrorAction Stop
                $map[$src] = "$($new.PrimarySmtpAddress)"
                Write-Log 'Contact' $c.DisplayName 'Created' $c.ExternalEmailAddress
            }
            catch { Write-Log 'Contact' $c.DisplayName 'Problem' $_.Exception.Message }
        }
        else { $map[$src] = $c.ExternalEmailAddress; Write-Log 'Contact' $c.DisplayName 'WhatIf' "Would create $($c.ExternalEmailAddress)" }
    }
}

# ---------------------------------------------------------- distribution lists and members
if ($Include -contains 'DistributionGroups') {
    Write-Host 'Distribution lists' -ForegroundColor Cyan
    $groupRows = @(Import-Csv (Join-Path $ExportFolder 'DistributionGroups.csv'))
    foreach ($g in $groupRows) {
        $src = $g.PrimarySmtpAddress
        $ours = Get-Ours $src
        if ($ours) { $map[$src] = "$($ours.PrimarySmtpAddress)"; Write-Log 'List' $g.DisplayName 'Exists' "$($ours.PrimarySmtpAddress)"; continue }
        $alias = "$AliasPrefix$($g.Alias)"
        if (Get-Recipient -Identity $alias -ErrorAction SilentlyContinue) {
            Write-Log 'List' $g.DisplayName 'Skipped' "Alias '$alias' is already used in the destination. Re-run with -AliasPrefix"
            continue
        }
        $address = "$alias@$AddressDomain"
        if ($PSCmdlet.ShouldProcess($address, 'Create distribution list')) {
            try {
                $type = if ($g.Security -eq 'True') { 'Security' } else { 'Distribution' }
                $name = $g.DisplayName
                if (Get-Recipient -Identity $name -ErrorAction SilentlyContinue) { $name = "$name ($alias)" }
                $new = New-DistributionGroup -Name $name -DisplayName $g.DisplayName -Alias $alias -PrimarySmtpAddress $address -Type $type -ErrorAction Stop
                $owners = @($g.Owners -split ';' | ForEach-Object { Get-Mapped $_ } | Where-Object { $_ })
                $set = @{
                    Identity                           = $new.Identity
                    CustomAttribute15                  = (Get-Tag $src)
                    RequireSenderAuthenticationEnabled = [bool]::Parse("$($g.RequireSenderAuthenticationEnabled)")
                    HiddenFromAddressListsEnabled      = [bool]::Parse("$($g.HiddenFromAddressListsEnabled)")
                    BypassSecurityGroupManagerCheck    = $true
                }
                if ($owners) { $set.ManagedBy = $owners }
                Set-DistributionGroup @set -ErrorAction Stop
                $map[$src] = $address
                Write-Log 'List' $g.DisplayName 'Created' "$address, $(@($owners).Count) owner(s)"
            }
            catch { Write-Log 'List' $g.DisplayName 'Problem' $_.Exception.Message }
        }
        else { $map[$src] = $address; Write-Log 'List' $g.DisplayName 'WhatIf' "Would create $address" }
    }

    Write-Host 'List members' -ForegroundColor Cyan
    $current = @{}
    foreach ($row in Import-Csv (Join-Path $ExportFolder 'DistributionGroupMembers.csv')) {
        $list = Get-Mapped $row.Group
        if (-not $list) { continue }   # the list itself was skipped or failed, already logged
        if (-not $row.Member) { Write-Log 'Member' "$($row.MemberName) -> $list" 'Skipped' "Has no email address in the source ($($row.MemberType)), so it can't be a list member here"; continue }
        $member = Get-Mapped $row.Member
        if (-not $member) { Write-Log 'Member' "$($row.Member) -> $list" 'Skipped' 'No mapping for this member'; continue }
        if (-not $WhatIfPreference -and -not $current.ContainsKey($list)) {
            $current[$list] = @(Get-DistributionGroupMember -Identity $list -ResultSize Unlimited -ErrorAction SilentlyContinue |
                ForEach-Object { "$($_.PrimarySmtpAddress)".ToLower() })
        }
        if ($current[$list] -contains $member.ToLower()) { continue }
        if ($PSCmdlet.ShouldProcess($list, "Add member $member")) {
            try {
                Add-DistributionGroupMember -Identity $list -Member $member -BypassSecurityGroupManagerCheck -ErrorAction Stop
                Write-Log 'Member' "$member -> $list" 'Added' ''
            }
            catch {
                if ($_.Exception.Message -match 'already a member') { Write-Log 'Member' "$member -> $list" 'Exists' '' }
                else { Write-Log 'Member' "$member -> $list" 'Problem' $_.Exception.Message }
            }
        }
        else { Write-Log 'Member' "$member -> $list" 'WhatIf' 'Would add' }
    }
}

# ---------------------------------------------------------- permissions
if ($Include -contains 'Permissions') {
    Write-Host 'Mailbox permissions' -ForegroundColor Cyan
    foreach ($p in Import-Csv (Join-Path $ExportFolder 'Permissions.csv')) {
        $box = Get-Mapped $p.Mailbox
        $who = Get-Mapped $p.Trustee
        $label = "$($p.Right): $($p.Trustee) on $($p.Mailbox)"
        if ($p.Trustee -like 'unresolved:*') { Write-Log 'Permission' $label 'Skipped' 'Trustee no longer exists in the source'; continue }
        if (-not $box -or -not $who) { Write-Log 'Permission' $label 'Skipped' 'Mailbox or trustee not in the mapping'; continue }
        if (-not $PSCmdlet.ShouldProcess($box, "$($p.Right) for $who")) { Write-Log 'Permission' $label 'WhatIf' "Would grant to $who on $box"; continue }
        try {
            switch ($p.Right) {
                'FullAccess'   { Add-MailboxPermission -Identity $box -User $who -AccessRights FullAccess -InheritanceType All -AutoMapping $true -WarningAction SilentlyContinue -ErrorAction Stop | Out-Null }
                'SendAs'       { Add-RecipientPermission -Identity $box -Trustee $who -AccessRights SendAs -Confirm:$false -WarningAction SilentlyContinue -ErrorAction Stop | Out-Null }
                'SendOnBehalf' { Set-Mailbox -Identity $box -GrantSendOnBehalfTo @{ Add = $who } -ErrorAction Stop }
            }
            Write-Log 'Permission' $label 'Added' "$who on $box"
        }
        catch {
            if ($_.Exception.Message -match 'already|exists') { Write-Log 'Permission' $label 'Exists' '' }
            else { Write-Log 'Permission' $label 'Problem' $_.Exception.Message }
        }
    }
}

# ---------------------------------------------------------- calendar and folder permissions
if ($Include -contains 'FolderPermissions') {
    Write-Host 'Calendar and folder permissions' -ForegroundColor Cyan
    $fpFile = Join-Path $ExportFolder 'FolderPermissions.csv'
    if (-not (Test-Path $fpFile)) {
        Write-Log 'FolderPermission' $ExportFolder 'Skipped' 'No FolderPermissions.csv. The export was made before 1.2.0 or with -SkipFolderPermissions. Re-run -Export'
    }
    else {
        foreach ($row in @(Import-Csv $fpFile)) {
            $label = "$($row.Folder): $($row.Trustee) on $($row.Mailbox)"
            $box = Get-Mapped $row.Mailbox
            if (-not $box) { Write-Log 'FolderPermission' $label 'Skipped' 'Mailbox not in the mapping'; continue }
            if ($row.Trustee -like 'external:*') { Write-Log 'FolderPermission' $label 'Skipped' 'Shared with someone outside the organisation. The owner re-shares it from the destination'; continue }
            if ($row.Trustee -like 'unresolved:*') { Write-Log 'FolderPermission' $label 'Skipped' 'Trustee no longer exists in the source'; continue }
            $isDefault = $row.Trustee -in 'Default', 'Anonymous'
            $who = if ($isDefault) { $row.Trustee } else { Get-Mapped $row.Trustee }
            if (-not $who) { Write-Log 'FolderPermission' $label 'Skipped' 'Trustee not in the mapping'; continue }
            $path = Get-DefaultFolderPath $box $row.Folder
            if (-not $path) {
                if ($WhatIfPreference) { $path = "\$($row.Folder)" }
                else { Write-Log 'FolderPermission' $label 'Problem' "No $($row.Folder) folder found in $box. Is the mailbox there yet?"; continue }
            }
            $fid = "${box}:$path"
            $rights = @($row.AccessRights -split ';' | Where-Object { $_ })
            $fp = @{ Identity = $fid; User = $who; AccessRights = $rights; ErrorAction = 'Stop' }
            # Delegate flags only apply to Editor on a calendar.
            if ($row.SharingPermissionFlags -and $row.Folder -eq 'Calendar' -and $rights -contains 'Editor') {
                $fp.SharingPermissionFlags = @($row.SharingPermissionFlags -split '[,;]\s*' | Where-Object { $_ })
            }
            if (-not $PSCmdlet.ShouldProcess($fid, "$($rights -join ',') for $who")) { Write-Log 'FolderPermission' $label 'WhatIf' "Would grant $($rights -join ',') to $who on $fid"; continue }
            try {
                $existing = $null
                if (-not $isDefault) {
                    try { $existing = Get-MailboxFolderPermission -Identity $fid -User $who -ErrorAction Stop } catch { $existing = $null }
                }
                if ($isDefault -or $existing) {
                    $now = if ($existing) { (@($existing.AccessRights) | ForEach-Object { "$_" }) -join ';' } else { '' }
                    if ($existing -and $now -eq ($rights -join ';')) { Write-Log 'FolderPermission' $label 'Exists' "$who already has $now"; continue }
                    Set-MailboxFolderPermission @fp | Out-Null
                    Write-Log 'FolderPermission' $label 'Set' "$who now $($rights -join ',') on $fid"
                }
                else {
                    Add-MailboxFolderPermission @fp | Out-Null
                    Write-Log 'FolderPermission' $label 'Added' "$who $($rights -join ',') on $fid$(if ($fp.SharingPermissionFlags) { " ($($fp.SharingPermissionFlags -join ', '))" })"
                }
            }
            catch {
                if ($_.Exception.Message -match 'already|exists') { Write-Log 'FolderPermission' $label 'Exists' '' }
                else { Write-Log 'FolderPermission' $label 'Problem' $_.Exception.Message }
            }
        }
    }
}

# ---------------------------------------------------------- forwarding
if ($Include -contains 'Forwarding') {
    Write-Host 'Forwarding' -ForegroundColor Cyan
    foreach ($f in Import-Csv (Join-Path $ExportFolder 'Forwarding.csv')) {
        $box = Get-Mapped $f.Mailbox
        $label = "$($f.Mailbox) -> $($f.ForwardTo)"
        if (-not $box) { Write-Log 'Forward' $label 'Skipped' 'Mailbox not in the mapping'; continue }
        $keep = [bool]::Parse("$($f.KeepCopy)")
        if ($f.Kind -eq 'Internal') {
            $to = Get-Mapped $f.ForwardTo
            if (-not $to) { Write-Log 'Forward' $label 'Skipped' 'Forwarding target not in the mapping'; continue }
            if ($PSCmdlet.ShouldProcess($box, "Forward to $to")) {
                try { Set-Mailbox -Identity $box -ForwardingAddress $to -DeliverToMailboxAndForward $keep -ErrorAction Stop; Write-Log 'Forward' $label 'Set' "$box -> $to, keep copy $keep" }
                catch { Write-Log 'Forward' $label 'Problem' $_.Exception.Message }
            }
            else { Write-Log 'Forward' $label 'WhatIf' "Would forward $box to $to" }
        }
        elseif ($IncludeExternalForwarding) {
            if ($PSCmdlet.ShouldProcess($box, "Forward externally to $($f.ForwardTo)")) {
                try { Set-Mailbox -Identity $box -ForwardingSmtpAddress $f.ForwardTo -DeliverToMailboxAndForward $keep -ErrorAction Stop; Write-Log 'Forward' $label 'Set' 'External. Check outbound spam policy allows it' }
                catch { Write-Log 'Forward' $label 'Problem' $_.Exception.Message }
            }
            else { Write-Log 'Forward' $label 'WhatIf' "Would forward $box externally" }
        }
        else { Write-Log 'Forward' $label 'Skipped' 'External forwarding. Use -IncludeExternalForwarding if it should stay' }
    }
}

# ---------------------------------------------------------- report
$file = ".\MailConfig-Import-$tenantName-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
$log | Export-Csv $file -NoTypeInformation -WhatIf:$false
Write-Host ''
$log | Group-Object Status | Sort-Object Name | ForEach-Object { Write-Host ("{0,-8} {1}" -f $_.Name, $_.Count) }
Write-Host "Report: $file" -ForegroundColor Cyan
if (@($log | Where-Object Status -eq 'Problem').Count) { Write-Host 'Some items failed. Fix them and re-run; anything already done is skipped.' -ForegroundColor Red }
Write-Host 'Transport rules are not imported. Rebuild them from the New-TransportRule commands in TransportRules.xml, disabled until tested.' -ForegroundColor Yellow
if (Test-Path (Join-Path $ExportFolder 'CheckAfterCutover.csv')) { Write-Host "After cutover, check inbox rules and automatic replies against $(Join-Path $ExportFolder 'CheckAfterCutover.csv')." -ForegroundColor Yellow }
