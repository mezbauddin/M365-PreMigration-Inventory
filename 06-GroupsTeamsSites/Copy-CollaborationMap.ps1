<#
.SYNOPSIS
    Maps every Microsoft 365 group, Team and SharePoint site in the source tenant, then creates the groups and Teams in
    the destination with their owners and members, ready for your copy tool to fill.

.DESCRIPTION
    TL;DR: Copy tools are happy to create groups and Teams for you, but then you get whatever address and site URL
    they pick. Run -Export against the source to get every group and Team with its addresses, visibility, site,
    storage, owners and members, plus the SharePoint sites that don't belong to a group. Run -Import against the
    destination to create the missing groups on your chosen domain, optionally turn them into Teams, and add owners
    and members mapped through your mapping CSV (Source,Destination). Then point your copy tool at the groups that
    already exist. Import is safe to re-run. Everything supports -WhatIf.

    -Export writes a folder of CSVs:
      Groups.csv       Every Microsoft 365 group: name, alias, addresses, visibility, Teams-enabled, site URL, storage,
                       owner and member counts
      GroupLinks.csv   One row per owner or member, by primary address
      Sites.csv        SharePoint sites that aren't group sites or OneDrives (needs -SharePointAdminUrl)
    -Import creates each missing group, sets its tag and turns welcome messages off, enables Teams if asked, then
    adds members before owners.

.PARAMETER Export
    Read the source tenant and write the export folder. Changes nothing.

.PARAMETER Import
    Create groups and add owners and members in the destination from an export folder.

.PARAMETER TenantId
    Optional tenant ID (GUID). The script stops if Exchange Online (or Graph, with -EnableTeams) is signed in to any
    other tenant.

.PARAMETER AdminUpn
    Optional admin sign-in name, e.g. admin@contoso.com. If Exchange Online is already signed in as someone else, or
    to another tenant, the script disconnects and signs in again as this account. Use it whenever you have both
    tenants open on one PC.

.PARAMETER SharePointAdminUrl
    With -Export: the source SharePoint admin URL, e.g. https://fabrikam-admin.sharepoint.com. Adds site storage and
    the list of non-group sites. Without it, Groups.csv has no storage figures and Sites.csv isn't written.

.PARAMETER OutputFolder
    With -Export: where to write. Defaults to .\Collaboration-<domain>-<date>.

.PARAMETER ExportFolder
    With -Import: the folder -Export wrote.

.PARAMETER MappingFile
    With -Import: CSV with Source and Destination columns, mapping each source address to its destination identity.
    The same file Copy-MailConfig and Switch-TargetAddresses use. A row for a group means 'use this existing
    destination group instead of creating one'.

.PARAMETER AddressDomain
    With -Import: accepted domain in the destination for the primary address of new groups, e.g. a staging domain.
    Source domains are added at cutover.

.PARAMETER AliasPrefix
    With -Import: prefix for new group aliases, for when the destination already uses names like 'sales'. The alias
    usually shapes the site URL too, so pick it with that in mind.

.PARAMETER EnableTeams
    With -Import: turn new and existing tagged groups into Teams where the source group was Teams-enabled, through
    Microsoft Graph (PUT /groups/{id}/team). Without it, groups are created as plain Microsoft 365 groups.

.PARAMETER Group
    With -Import: only these source group addresses. Handy for doing one Team at a time.

.EXAMPLE
    .\Copy-CollaborationMap.ps1 -Export -TenantId <source-guid> -SharePointAdminUrl https://fabrikam-admin.sharepoint.com

.EXAMPLE
    .\Copy-CollaborationMap.ps1 -Import -TenantId <dest-guid> -ExportFolder .\Collaboration-fabrikam.com-20261109-1100 `
        -MappingFile .\UserMapping.csv -AddressDomain contoso.com -EnableTeams -WhatIf

.EXAMPLE
    .\Copy-CollaborationMap.ps1 -Import -TenantId <dest-guid> -ExportFolder .\Collaboration-fabrikam.com-20261109-1100 `
        -MappingFile .\UserMapping.csv -AddressDomain contoso.com -Group sales@fabrikam.com

.EXAMPLE
    .\Copy-CollaborationMap.ps1 -Import -TenantId <dest-guid> -AdminUpn admin@contoso.com `
        -ExportFolder .\Collaboration-fabrikam.com-20261109-1100 -MappingFile .\UserMapping.csv -AddressDomain contoso.com

.NOTES
    Author  : Mezba Uddin (Microsoft MVP)
    Version : 1.1.0
    Date    : 2026-11-10
    Licence : MIT
    Repo    : https://github.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit
    Requires: ExchangeOnlineManagement 3.x, Exchange Administrator. For -SharePointAdminUrl:
              Microsoft.Online.SharePoint.PowerShell and SharePoint Administrator. For -EnableTeams:
              Microsoft.Graph.Authentication with Group.ReadWrite.All.
    Groups this script creates are tagged with CustomAttribute15 = 'T2T:<source address>', so re-runs skip them and
    Switch-TargetAddresses (part 7) gives them their old address at cutover.
    The signed-in admin is the first owner of every group it creates, and stays an owner. Remove yourself afterwards
    with Remove-UnifiedGroupLinks if you don't want to be.
    Private and shared channels are not created. Sites.csv lists their sites so you know they exist.
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Export')]
param(
    [Parameter(ParameterSetName = 'Export', Mandatory)][switch]$Export,
    [Parameter(ParameterSetName = 'Import', Mandatory)][switch]$Import,
    [ValidatePattern('^[0-9a-fA-F-]{36}$')][string]$TenantId,
    [string]$AdminUpn,

    [Parameter(ParameterSetName = 'Export')][string]$SharePointAdminUrl,
    [Parameter(ParameterSetName = 'Export')][string]$OutputFolder,

    [Parameter(ParameterSetName = 'Import', Mandatory)][string]$ExportFolder,
    [Parameter(ParameterSetName = 'Import', Mandatory)][string]$MappingFile,
    [Parameter(ParameterSetName = 'Import', Mandatory)][string]$AddressDomain,
    [Parameter(ParameterSetName = 'Import')][string]$AliasPrefix = '',
    [Parameter(ParameterSetName = 'Import')][switch]$EnableTeams,
    [Parameter(ParameterSetName = 'Import')][string[]]$Group = @()
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
Write-Host "Connected to $tenantName as $($conn.UserPrincipalName)" -ForegroundColor Cyan

function Get-SmtpOnly($a) { if ($a) { ("$a" -replace '^smtp:', '').Trim().ToLower() } }
function Get-Tag($src) { "T2T:$($src.ToLower())" }
function Get-LinkAddress($l) {
    if ($l.PrimarySmtpAddress) { "$($l.PrimarySmtpAddress)".ToLower() } else { "unresolved:$($l.Name)" }
}

# ============================================================ EXPORT
if ($Export) {
    if (-not $OutputFolder) { $OutputFolder = ".\Collaboration-$tenantName-$(Get-Date -Format 'yyyyMMdd-HHmm')" }
    New-Item -ItemType Directory -Path $OutputFolder -Force -WhatIf:$false | Out-Null

    # ---------------------------------------------------- SharePoint, optional
    $sites = @()
    if ($SharePointAdminUrl) {
        if ($SharePointAdminUrl -notmatch '^https://') { $SharePointAdminUrl = "https://$SharePointAdminUrl" }
        try {
            if (-not (Get-Command Connect-SPOService -ErrorAction SilentlyContinue)) {
                try { Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -ErrorAction Stop }
                catch {
                    if ($PSVersionTable.PSEdition -eq 'Core') {
                        Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -DisableNameChecking -WarningAction SilentlyContinue
                    }
                    else { throw }
                }
            }
            Connect-SPOService -Url $SharePointAdminUrl
            $sites = @(Get-SPOSite -Limit All)
            Write-Host "$($sites.Count) SharePoint sites read." -ForegroundColor Cyan
        }
        catch {
            Write-Warning "Couldn't read SharePoint ($($_.Exception.Message)). Carrying on without storage figures or Sites.csv."
            $sites = @()
        }
    }
    else { Write-Warning 'No -SharePointAdminUrl, so no storage figures and no list of non-group sites.' }
    $siteByUrl = @{}
    foreach ($s in $sites) { $siteByUrl["$($s.Url)".TrimEnd('/').ToLower()] = $s }

    # ---------------------------------------------------- groups and their links
    $groups = @(Get-UnifiedGroup -ResultSize Unlimited)
    Write-Host "$($groups.Count) Microsoft 365 groups. Reading owners and members." -ForegroundColor Cyan
    $links = [System.Collections.Generic.List[object]]::new()
    $i = 0
    $groupRows = foreach ($g in $groups) {
        $i++; Write-Progress -Activity 'Groups' -Status $g.DisplayName -PercentComplete (100 * $i / $groups.Count)
        $gAddr = "$($g.PrimarySmtpAddress)".ToLower()
        $id = $g.ExternalDirectoryObjectId
        $owners = @(Get-UnifiedGroupLinks -Identity $id -LinkType Owners -ResultSize Unlimited -ErrorAction SilentlyContinue | ForEach-Object { Get-LinkAddress $_ })
        $members = @(Get-UnifiedGroupLinks -Identity $id -LinkType Members -ResultSize Unlimited -ErrorAction SilentlyContinue | ForEach-Object { Get-LinkAddress $_ })
        foreach ($o in $owners) { $links.Add([pscustomobject]@{ Group = $gAddr; Address = $o; Role = 'Owner' }) }
        foreach ($m in $members) { $links.Add([pscustomobject]@{ Group = $gAddr; Address = $m; Role = 'Member' }) }

        $siteUrl = "$($g.SharePointSiteUrl)".TrimEnd('/')
        $site = if ($siteUrl) { $siteByUrl[$siteUrl.ToLower()] }
        [pscustomobject]@{
            DisplayName                   = $g.DisplayName
            Alias                         = $g.Alias
            PrimarySmtpAddress            = $gAddr
            AllSmtpAddresses              = (@($g.EmailAddresses | Where-Object { $_ -like 'smtp:*' } | ForEach-Object { Get-SmtpOnly $_ } | Select-Object -Unique) -join ';')
            Visibility                    = $g.AccessType
            TeamsEnabled                  = (@($g.ResourceProvisioningOptions) -contains 'Team')
            SharePointSiteUrl             = $siteUrl
            StorageGB                     = if ($site) { [math]::Round($site.StorageUsageCurrent / 1024, 2) } else { $null }
            OwnerCount                    = $owners.Count
            MemberCount                   = $members.Count
            Description                   = $g.Notes
            HiddenFromAddressListsEnabled = $g.HiddenFromAddressListsEnabled
        }
    }
    Write-Progress -Activity 'Groups' -Completed
    $groupRows | Export-Csv (Join-Path $OutputFolder 'Groups.csv') -NoTypeInformation -WhatIf:$false
    $links     | Export-Csv (Join-Path $OutputFolder 'GroupLinks.csv') -NoTypeInformation -WhatIf:$false

    # ---------------------------------------------------- sites that aren't group sites or OneDrives
    $siteRows = @()
    if ($sites) {
        $groupSites = @($groupRows | ForEach-Object { "$($_.SharePointSiteUrl)".ToLower() } | Where-Object { $_ })
        $system = 'SPSMSITEHOST#0', 'SRCHCEN#0', 'APPCATALOG#0', 'REDIRECTSITE#0', 'POINTPUBLISHINGHUB#0', 'POINTPUBLISHINGTOPIC#0', 'EDISC#0', 'SPSMSITEHOST#1'
        $siteRows = foreach ($s in $sites) {
            $url = "$($s.Url)".TrimEnd('/')
            if ($url -like '*-my.sharepoint.com*' -or $groupSites -contains $url.ToLower() -or "$($s.Template)" -like 'GROUP#*') { continue }
            $kind = if ("$($s.Template)" -like 'TEAMCHANNEL#*') { 'PrivateOrSharedChannel' }
                    elseif ($system -contains "$($s.Template)") { 'System' }
                    else { 'Site' }
            [pscustomobject]@{
                Url                     = $url
                Title                   = $s.Title
                Template                = $s.Template
                Kind                    = $kind
                Owner                   = $s.Owner
                StorageGB               = [math]::Round($s.StorageUsageCurrent / 1024, 2)
                LastContentModifiedDate = $s.LastContentModifiedDate
            }
        }
        $siteRows | Sort-Object Kind, Url | Export-Csv (Join-Path $OutputFolder 'Sites.csv') -NoTypeInformation -WhatIf:$false
    }

    $teams = @($groupRows | Where-Object TeamsEnabled).Count
    $channels = @($siteRows | Where-Object Kind -eq 'PrivateOrSharedChannel').Count
    $unresolved = @($links | Where-Object Address -like 'unresolved:*').Count
    Write-Host ''
    Write-Host "Export written to $(Resolve-Path $OutputFolder)" -ForegroundColor Green
    Write-Host ("  {0} groups ({1} Teams), {2} owner and member links, {3} other sites" -f
        @($groupRows).Count, $teams, $links.Count, @($siteRows | Where-Object Kind -eq 'Site').Count)
    if ($channels) { Write-Host "  $channels private or shared channel site(s). Import doesn't create these channels. Plan them by hand." -ForegroundColor Yellow }
    if ($unresolved) { Write-Host "  $unresolved link(s) have no mail address (Address 'unresolved:'). Import skips them." -ForegroundColor Yellow }
    Write-Host 'The CSVs hold names and addresses. Keep them out of source control and shared drives.' -ForegroundColor Yellow
    return
}

# ============================================================ IMPORT
if ($accepted -notcontains $AddressDomain.ToLower()) { throw "$AddressDomain isn't an accepted domain in $tenantName." }
foreach ($f in 'Groups.csv', 'GroupLinks.csv') {
    if (-not (Test-Path (Join-Path $ExportFolder $f))) { throw "$f missing from $ExportFolder. Point -ExportFolder at the folder -Export wrote." }
}
$admin = "$($conn.UserPrincipalName)"
if (-not $admin) { throw "Couldn't read the signed-in account from Get-ConnectionInformation. It's needed as the first owner of new groups." }

if ($EnableTeams) {
    $graphTenant = if ($TenantId) { $TenantId } else { $conn.TenantID }
    $mg = Get-MgContext
    if ($mg -and (($mg.TenantId -ne $graphTenant) -or ($AdminUpn -and $mg.Account -ne $AdminUpn))) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    Connect-MgGraph -TenantId $graphTenant -NoWelcome -Scopes 'Group.ReadWrite.All'
    if ((Get-MgContext).TenantId -ne $graphTenant) { throw "Graph signed in to $((Get-MgContext).TenantId), not $graphTenant." }
    if ($AdminUpn -and (Get-MgContext).Account -ne $AdminUpn) { throw "Graph signed in as $((Get-MgContext).Account), not $AdminUpn. Run Disconnect-MgGraph and pick the right account." }
    Write-Host "Graph signed in as $((Get-MgContext).Account) to tenant $graphTenant" -ForegroundColor Cyan
}

$map = @{}
Import-Csv $MappingFile | ForEach-Object {
    if (-not $_.Source -or -not $_.Destination) { throw "$MappingFile needs Source and Destination columns." }
    $map[$_.Source.Trim().ToLower()] = $_.Destination.Trim()
}
Write-Host "$($map.Count) mappings loaded." -ForegroundColor Cyan

$log = [System.Collections.Generic.List[object]]::new()
function Write-Log($Area, $Object, $Status, $Detail) {
    $log.Add([pscustomobject]@{ Area = $Area; Object = $Object; Status = $Status; Detail = $Detail })
    $c = @{ Created = 'Green'; Added = 'Green'; Set = 'Green'; Exists = 'Gray'; Skipped = 'Yellow'; Problem = 'Red'; WhatIf = 'Cyan' }[$Status]
    Write-Host ("  [{0}] {1}: {2} {3}" -f $Status, $Area, $Object, $Detail) -ForegroundColor $c
}
function Get-Ours($src) { Get-Recipient -Filter "CustomAttribute15 -eq '$(Get-Tag $src)'" -ErrorAction SilentlyContinue | Select-Object -First 1 }

function Add-Team($groupId, $label) {
    # Graph can return 404 for a short while after a group is created, so try a few times.
    try { Invoke-MgGraphRequest -Method GET -Uri "v1.0/groups/$groupId/team" -ErrorAction Stop | Out-Null; Write-Log 'Team' $label 'Exists' ''; return }
    catch { if ($_.Exception.Message -notmatch '404|NotFound') { Write-Log 'Team' $label 'Problem' $_.Exception.Message; return } }
    if (-not $PSCmdlet.ShouldProcess($label, 'Enable Teams')) { Write-Log 'Team' $label 'WhatIf' 'Would enable Teams'; return }
    for ($try = 1; $try -le 3; $try++) {
        try {
            Invoke-MgGraphRequest -Method PUT -Uri "v1.0/groups/$groupId/team" -Body '{}' -ContentType 'application/json' -ErrorAction Stop | Out-Null
            Write-Log 'Team' $label 'Created' ''
            return
        }
        catch {
            if ($try -lt 3 -and $_.Exception.Message -match '404|NotFound') { Start-Sleep -Seconds 10; continue }
            Write-Log 'Team' $label 'Problem' "$($_.Exception.Message). Re-run later; the group is fine"
            return
        }
    }
}

$groupRows = @(Import-Csv (Join-Path $ExportFolder 'Groups.csv'))
$links = @(Import-Csv (Join-Path $ExportFolder 'GroupLinks.csv'))
$only = @($Group | ForEach-Object { $_.Trim().ToLower() })
if ($only) {
    $groupRows = @($groupRows | Where-Object { $only -contains $_.PrimarySmtpAddress.ToLower() })
    if (-not $groupRows) { throw "None of $($only -join ', ') are in Groups.csv." }
}
Write-Host "$($groupRows.Count) group(s) to process." -ForegroundColor Cyan

foreach ($g in $groupRows) {
    $src = $g.PrimarySmtpAddress.ToLower()
    Write-Host $g.DisplayName -ForegroundColor Cyan

    # ---------------------------------------------------- find or create the group
    $target = $null; $targetId = $null
    $ours = Get-Ours $src
    if ($ours) {
        $target = "$($ours.PrimarySmtpAddress)"; $targetId = $ours.ExternalDirectoryObjectId
        Write-Log 'Group' $g.DisplayName 'Exists' $target
    }
    elseif ($map.ContainsKey($src)) {
        $mapped = Get-Recipient -Identity $map[$src] -ErrorAction SilentlyContinue
        if (-not $mapped) { Write-Log 'Group' $g.DisplayName 'Skipped' "Mapping points at $($map[$src]), which doesn't exist in the destination"; continue }
        if ("$($mapped.RecipientTypeDetails)" -ne 'GroupMailbox') { Write-Log 'Group' $g.DisplayName 'Skipped' "Mapping points at $($map[$src]), which is a $($mapped.RecipientTypeDetails), not a Microsoft 365 group"; continue }
        $target = "$($mapped.PrimarySmtpAddress)"; $targetId = $mapped.ExternalDirectoryObjectId
        Write-Log 'Group' $g.DisplayName 'Exists' "Using $target from the mapping file"
    }
    else {
        $alias = "$AliasPrefix$($g.Alias)"
        $address = "$alias@$AddressDomain"
        if ((Get-Recipient -Identity $alias -ErrorAction SilentlyContinue) -or (Get-Recipient -Identity $address -ErrorAction SilentlyContinue)) {
            Write-Log 'Group' $g.DisplayName 'Skipped' "Alias '$alias' is already used in the destination. Re-run with -AliasPrefix, or map it to an existing group"
            continue
        }
        if ($PSCmdlet.ShouldProcess($address, 'Create Microsoft 365 group')) {
            try {
                $access = if ($g.Visibility -eq 'Public') { 'Public' } else { 'Private' }
                $new = @{ DisplayName = $g.DisplayName; Alias = $alias; PrimarySmtpAddress = $address; AccessType = $access; Owner = $admin; ErrorAction = 'Stop' }
                if ($g.Description) { $new.Notes = $g.Description }
                $created = New-UnifiedGroup @new
                $targetId = $created.ExternalDirectoryObjectId
                Set-UnifiedGroup -Identity $targetId -CustomAttribute15 (Get-Tag $src) -UnifiedGroupWelcomeMessageEnabled:$false `
                    -HiddenFromAddressListsEnabled ("$($g.HiddenFromAddressListsEnabled)" -eq 'True') -ErrorAction Stop
                $target = $address
                Write-Log 'Group' $g.DisplayName 'Created' "$address, $access, welcome messages off"
            }
            catch { Write-Log 'Group' $g.DisplayName 'Problem' $_.Exception.Message; continue }
        }
        else { $target = $address; Write-Log 'Group' $g.DisplayName 'WhatIf' "Would create $address ($($g.Visibility)), owner $admin" }
    }

    # ---------------------------------------------------- Teams
    if ($g.TeamsEnabled -eq 'True') {
        if ($EnableTeams) {
            if ($targetId) { Add-Team $targetId $g.DisplayName }
            else { Write-Log 'Team' $g.DisplayName 'WhatIf' 'Would enable Teams once the group exists' }
        }
        else { Write-Log 'Team' $g.DisplayName 'Skipped' 'Teams-enabled in the source. Use -EnableTeams, or let your copy tool add the team' }
    }

    # ---------------------------------------------------- members, then owners
    $current = @{ Members = @(); Owners = @() }
    if ($targetId) {
        foreach ($t in 'Members', 'Owners') {
            $current[$t] = @(Get-UnifiedGroupLinks -Identity $targetId -LinkType $t -ResultSize Unlimited -ErrorAction SilentlyContinue |
                ForEach-Object { "$($_.PrimarySmtpAddress)".ToLower() })
        }
    }
    $mine = @($links | Where-Object { $_.Group.ToLower() -eq $src })
    $wantOwners = @($mine | Where-Object Role -eq 'Owner' | ForEach-Object { $_.Address.ToLower() })
    # Owners must be members first. Add any owner who isn't already a member in the source list.
    $wantMembers = @(@($mine | Where-Object Role -eq 'Member' | ForEach-Object { $_.Address.ToLower() }) + $wantOwners | Select-Object -Unique)

    foreach ($pass in @(@{ Type = 'Members'; Want = $wantMembers; Label = 'Member' }, @{ Type = 'Owners'; Want = $wantOwners; Label = 'Owner' })) {
        foreach ($who in $pass.Want) {
            $label = "$who -> $target"
            if ($who -like 'unresolved:*') { Write-Log $pass.Label $label 'Skipped' 'No mail address in the source'; continue }
            $dest = $map[$who]
            if (-not $dest) { Write-Log $pass.Label $label 'Skipped' 'Not in the mapping'; continue }
            if ($current[$pass.Type] -contains $dest.ToLower()) { continue }
            if (-not $targetId -or -not $PSCmdlet.ShouldProcess($target, "Add $($pass.Label.ToLower()) $dest")) {
                Write-Log $pass.Label "$dest -> $target" 'WhatIf' 'Would add'; continue
            }
            try {
                Add-UnifiedGroupLinks -Identity $targetId -LinkType $pass.Type -Links $dest -Confirm:$false -ErrorAction Stop
                Write-Log $pass.Label "$dest -> $target" 'Added' ''
            }
            catch {
                if ($_.Exception.Message -match 'already') { Write-Log $pass.Label "$dest -> $target" 'Exists' '' }
                else { Write-Log $pass.Label "$dest -> $target" 'Problem' $_.Exception.Message }
            }
        }
    }
}

# ---------------------------------------------------------- report
$file = ".\Collaboration-Import-$tenantName-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
$log | Export-Csv $file -NoTypeInformation -WhatIf:$false
Write-Host ''
$log | Group-Object Status | Sort-Object Name | ForEach-Object { Write-Host ("{0,-8} {1}" -f $_.Name, $_.Count) }
Write-Host "Report: $file" -ForegroundColor Cyan
if (@($log | Where-Object Status -eq 'Problem').Count) { Write-Host 'Some items failed. Fix them and re-run; anything already done is skipped.' -ForegroundColor Red }
Write-Host "You ($admin) are an owner of every group this created. Remove yourself with Remove-UnifiedGroupLinks once the copy is done, if you don't want to be." -ForegroundColor Yellow
Write-Host 'Private and shared channels are not created. Import the same user mapping into your copy tool before you start copying.' -ForegroundColor Yellow
