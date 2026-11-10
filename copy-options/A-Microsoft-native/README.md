# Option A: Microsoft native cross-tenant migration

**TL;DR:** Microsoft can move mailboxes (with archives) and OneDrive between tenants itself, server to server, keeping items, folders and permissions. You pay a one-time add-on per user, set up a trust between the tenants, and build each destination person as a **mail user** stamped with the source mailbox's identity, not as a mailbox. This page is the full command list, in order, ready to paste, plus [`New-TargetMailUsers.ps1`](New-TargetMailUsers.ps1), which does the stamping (the step people get wrong). It doesn't move shared Teams content, and the OneDrive move is a one-time move, not a copy with incrementals.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 2.0.0 |
| Status | Every step below is from Microsoft's documentation. None is lab-verified yet; the Verified column changes as I run each one end to end |
| Script | [`New-TargetMailUsers.ps1`](New-TargetMailUsers.ps1) 1.0.0 |
| Licence | MIT |

Part of the [copy options](../README.md) in my [M365 tenant-to-tenant migration toolkit](../../README.md).

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 2.0.0 | 2026-11-10 | Pasteable command list for mailbox and OneDrive moves, licensing order, lab-verified vs docs marking. Ships `New-TargetMailUsers.ps1` 1.0.0. Removes the list of planned scripts that didn't exist |
| 1.0.0 | 2026-11-09 | First version: the plan and how it differs from the rest of the toolkit |

## Why use it

- **Nothing to install, nothing in the middle.** Mail moves on Microsoft's own move service, the same one hybrid moves use. No EWS, so [03 EWS archive readiness](../../03-EwsArchiveReadiness/) doesn't apply and the EWS shutdown on 1 April 2027 doesn't touch it.
- **It's a real move.** The mailbox keeps its ExchangeGuid, so Outlook profiles, permissions held inside the mailbox, and the archive come across as they were.
- **Big archives.** Server to server copes better with a 100 GB archive than anything copying item by item.

## What happens without it

Native isn't compulsory; B and C exist for a reason. What you lose by not using it:

- **Archives go over EWS** with most third-party tools. That needs the EWS allow list in both tenants ([03](../../03-EwsArchiveReadiness/)), and it stops working on 1 April 2027.
- **OneDrive links break.** A native OneDrive move leaves a redirect at the old URL. A copy doesn't.

And what goes wrong if you use native but skip the steps below: the move for that person fails, usually with a message about a missing ExchangeGuid or X500, or the destination already holds an empty mailbox you then have to unpick.

## What it covers

| Workload | Native route | Licence |
| --- | --- | --- |
| Mailboxes and In-Place Archives | Cross-tenant mailbox migration (migration batches run from the destination) | Cross-Tenant User Data Migration add-on, per user, one-time |
| OneDrive | Cross-tenant OneDrive migration (SharePoint Online PowerShell) | Same add-on |
| SharePoint sites | Cross-tenant SharePoint migration | Cross-tenant shared data migration, per 100 GB. Enterprise Agreement only at the time of writing |
| Teams channels, chat in channels, Planner, Power Platform | Not covered | n/a |

Check Microsoft's current pricing and availability before you quote. Both changed during 2026.

## Licensing, and the order that matters

- **Cross-Tenant User Data Migration add-on:** one per person, one-time, assigned to the user in either tenant. It covers that person's mailbox and OneDrive. Microsoft's docs are blunt: no licence, the move fails, and there are no exceptions. Buy it a week early so it's there when you need it.
- **Trap: don't license the destination before the move.** Give an Exchange licence to a destination user that has no ExchangeGuid stamped and Exchange creates a fresh, empty mailbox. The move for that person then fails, and you're deleting a mailbox and re-stamping by hand. Microsoft's docs allow licensing once the mail user is stamped; their FAQ says it's safest to license after the move completes, inside the 30-day grace period. I go with the FAQ: **stamp, move, then license the mail user.**
- **Trap: OneDrive wants a licence first, and no OneDrive yet.** The OneDrive move needs the destination user licensed for OneDrive, with no OneDrive created. So either move mail first, license, then move OneDrive, or license with the Exchange Online service plan switched off until the mailbox is in. And don't run [04 Target OneDrives](../../04-TargetOneDrives/) for anyone moving this way: a pre-created OneDrive makes the move fail.

## How it differs from the rest of this toolkit

- **[02 Target mailboxes](../../02-TargetMailboxes/) creates mailboxes.** That's right for ShareGate and the no-budget route, and wrong here. Don't run `02 -Apply` for anyone moving natively. Use [`New-TargetMailUsers.ps1`](New-TargetMailUsers.ps1) instead.
- **The mapping still drives everything.** Build `Source,Destination` with [02b](../../02b-UserMapping/) as usual. `New-TargetMailUsers.ps1` reads it, and so do 05, 07 and 09.
- **The rest of the spine is unchanged:** [05 Mail config](../../05-MailConfig/), [06 Groups, Teams and sites](../../06-GroupsTeamsSites/), [07 Domain cutover](../../07-DomainCutover/), [09 Validation](../../09-Validation/).

## The commands, in order

Placeholders: `fabrikam` is the source, `contoso` the destination. Replace `<...>` values. **Verified** says whether I've run the step in my lab tenants (Lab) or it's straight from Microsoft's docs (Docs). Right now, all of it is Docs.

Every Exchange step: name the account (`Connect-ExchangeOnline -UserPrincipalName ...`) and check `Get-ConnectionInformation` shows the tenant you expect. With two tenants on one PC it's easy to run a source command in the destination.

### Mailboxes

| # | Where | Step | Verified |
| --- | --- | --- | --- |
| 1 | Both | Buy and assign the Cross-Tenant User Data Migration add-on | Docs |
| 2 | Destination | App registration with `Mailbox.Migration`, secret, admin consent | Docs |
| 3 | Source | Admin consent to the same app | Docs |
| 4 | Destination | Organisation relationship (Inbound) and migration endpoint | Docs |
| 5 | Source | Mail-enabled security group, organisation relationship (RemoteOutbound) | Docs |
| 6 | Destination | Mail users stamped from the source (`New-TargetMailUsers.ps1`) | Docs |
| 7 | Destination | Test batch, then waves | Docs |
| 8 | Destination | Complete batches, then license | Docs |

**1. Licences.** Microsoft 365 admin centre > Marketplace, search for Cross-Tenant User Data Migration, buy one per person moving, assign them.

**2. Destination: the migration app.** Entra admin centre in the **destination** > App registrations > New registration:

- Supported account types: **Accounts in any organisational directory** (multitenant).
- Redirect URI: Web, `https://office.com`.
- API permissions > Add > APIs my organisation uses > **Office 365 Exchange Online** > Application permissions > **Mailbox.Migration**. Grant admin consent.
- Certificates and secrets > New client secret. Copy the value now; you can't see it again. Note the Application (client) ID.

**3. Source: consent.** A source Global Administrator opens this in a browser and accepts:

```text
https://login.microsoftonline.com/fabrikam.onmicrosoft.com/adminconsent?client_id=<app-id>&redirect_uri=https://office.com
```

**Trap:** the endpoint in step 4 fails with an authentication error until the source has consented. Do step 3 first.

**4. Destination: organisation relationship and endpoint.**

```powershell
Connect-ExchangeOnline -UserPrincipalName admin@contoso.onmicrosoft.com
$sourceTenantId = '<source tenant ID>'
$appId  = '<app-id>'
$secret = Read-Host 'App secret' -AsSecureString

if ((Get-OrganizationConfig).IsDehydrated) { Enable-OrganizationCustomization }

$rel = Get-OrganizationRelationship | Where-Object { $_.DomainNames -like $sourceTenantId }
if ($rel) { Set-OrganizationRelationship $rel.Name -Enabled:$true -MailboxMoveEnabled:$true -MailboxMoveCapability Inbound }
else { New-OrganizationRelationship 'T2T from Fabrikam' -Enabled:$true -MailboxMoveEnabled:$true -MailboxMoveCapability Inbound -DomainNames $sourceTenantId }

$cred = [pscredential]::new($appId, $secret)
New-MigrationEndpoint -Name 'T2T-Fabrikam' -RemoteServer outlook.office.com -RemoteTenant fabrikam.onmicrosoft.com `
    -Credentials $cred -ApplicationId $appId -ExchangeRemoteMove:$true
```

Microsoft's current example passes the secret as `-Credentials` (app ID as the user name, secret as the password), as above. Some write-ups pass the secret with `-AppSecretKeyToSecure` instead; `Get-Help New-MigrationEndpoint -Parameter *Secret*` shows what your module version takes.

**5. Source: scope group and organisation relationship.**

```powershell
Connect-ExchangeOnline -UserPrincipalName admin@fabrikam.onmicrosoft.com
$targetTenantId = '<destination tenant ID>'
$appId = '<app-id>'

if ((Get-OrganizationConfig).IsDehydrated) { Enable-OrganizationCustomization }
New-DistributionGroup -Type Security -Name 'T2T-Movers' -PrimarySmtpAddress t2t-movers@fabrikam.com
Import-Csv .\UserMapping.csv | ForEach-Object { Add-DistributionGroupMember 'T2T-Movers' -Member $_.Source -ErrorAction Continue }

$rel = Get-OrganizationRelationship | Where-Object { $_.DomainNames -like $targetTenantId }
if ($rel) { Set-OrganizationRelationship $rel.Name -Enabled:$true -MailboxMoveEnabled:$true -MailboxMoveCapability RemoteOutbound -OAuthApplicationId $appId -MailboxMovePublishedScopes 'T2T-Movers' }
else { New-OrganizationRelationship 'T2T to Contoso' -Enabled:$true -MailboxMoveEnabled:$true -MailboxMoveCapability RemoteOutbound -DomainNames $targetTenantId -OAuthApplicationId $appId -MailboxMovePublishedScopes 'T2T-Movers' }
```

**Tip:** only mailboxes in the scope group can move. Add people directly rather than through nested groups, and check membership before each wave. A mailbox missing from the group fails its move with a scope error.

**6. Destination: stamped mail users.** Every person needs a destination mail user carrying:

| Attribute | Value |
| --- | --- |
| ExchangeGuid | The source mailbox's ExchangeGuid. No match, no move |
| ArchiveGuid | The source archive's GUID, only if the source archive is active |
| EmailAddresses | `x500:<source LegacyExchangeDN>`, plus every x500 the source mailbox already has |
| UserPrincipalName and primary SMTP | The new identity: the mapping's Destination, on a destination domain |
| ExternalEmailAddress | The source mailbox's `fabrikam.onmicrosoft.com` address. Set it together with the primary SMTP, or Exchange makes it the primary and the move fails |

[`New-TargetMailUsers.ps1`](New-TargetMailUsers.ps1) does all of that from the mapping file:

```powershell
# Easiest: pass 01's 04-Mailboxes.csv (inventory 1.1.0 or later) as -SourceCsv. Or read the source live. Changes nothing.
.\New-TargetMailUsers.ps1 -MappingFile .\UserMapping.csv -SourceAdminUpn admin@fabrikam.onmicrosoft.com -ExportOnly

# Preview, then stamp, a two-person test batch
.\New-TargetMailUsers.ps1 -MappingFile .\UserMapping.csv -SourceCsv .\SourceMailboxIdentity-<date>.csv `
    -AdminUpn admin@contoso.onmicrosoft.com -Only alex@fabrikam.com, kim@fabrikam.com -WhatIf
.\New-TargetMailUsers.ps1 -MappingFile .\UserMapping.csv -SourceCsv .\SourceMailboxIdentity-<date>.csv `
    -AdminUpn admin@contoso.onmicrosoft.com -Only alex@fabrikam.com, kim@fabrikam.com

# Everyone
.\New-TargetMailUsers.ps1 -MappingFile .\UserMapping.csv -SourceCsv .\SourceMailboxIdentity-<date>.csv -AdminUpn admin@contoso.onmicrosoft.com
```

What it does per person:

| Destination has | Action |
| --- | --- |
| Nothing | `New-MailUser`, then stamps it |
| A user with no mail attributes | `Enable-MailUser`, then stamps it |
| A mail user | Stamps it. If it already has a different ExchangeGuid, marks it Conflict unless you add `-Force` |
| A mailbox | Conflict. The move would fail. Remove the mailbox and its licence, or use another option for this person |

It writes `TargetMailUsers-Report-<date>.csv` (Create, Update, Enable, Skip, Conflict, Error, with the reason) and `MigrationBatch-<date>.csv` for step 7. It never assigns a licence. New mail users get a random password nobody knows; [07](../../07-DomainCutover/) sets the temporary password people use at cutover.

**Tip:** the source's own addresses (`alex@fabrikam.com`) can't go on the destination object yet, because the destination doesn't own `fabrikam.com` until cutover. [07 Switch-TargetAddresses](../../07-DomainCutover/) adds them then.

**7. Destination: test batch, then waves.**

```powershell
Connect-ExchangeOnline -UserPrincipalName admin@contoso.onmicrosoft.com
# CSV header is EmailAddress, and the address is the DESTINATION one (MigrationBatch-<date>.csv from step 6)
New-MigrationBatch -Name 'T2T-Test' -SourceEndpoint 'T2T-Fabrikam' -CSVData ([IO.File]::ReadAllBytes("$PWD\MigrationBatch-<date>.csv")) `
    -TargetDeliveryDomain contoso.onmicrosoft.com -Autostart
Get-MigrationBatch 'T2T-Test' | Format-List Status, TotalCount, SyncedCount, FailedCount
Get-MigrationUser -BatchId 'T2T-Test' | Get-MigrationUserStatistics | Format-Table Identity, Status, PercentageComplete, Error
```

Check the test people properly: item counts, archive, calendar, a shared mailbox's permissions. Then build a batch per wave.

**8. Complete, then license.** Batches sync and wait. Complete each just before cutover:

```powershell
Complete-MigrationBatch -Identity 'T2T-Wave1'
# or let it finish on its own at a set time (UTC)
Set-MigrationBatch -Identity 'T2T-Wave1' -CompleteAfter '<yyyy-MM-ddTHH:mm:ssZ>'
```

When a person shows Completed, the destination mail user is now a mailbox. License them, then check the type: if a shared mailbox arrived as a user mailbox, convert it (`Set-Mailbox <address> -Type Shared`). The source mailbox becomes a mail user pointing at the destination.

### OneDrive

Run in the SharePoint Online Management Shell (latest version), with `Connect-SPOService -Url https://fabrikam-admin.sharepoint.com` for the source and `https://contoso-admin.sharepoint.com` for the destination.

| # | Where | Step | Verified |
| --- | --- | --- | --- |
| 1 | Both | Get each tenant's cross-tenant host URL | Docs |
| 2 | Source, then destination | Trust each other | Docs |
| 3 | Both | Verify the trust | Docs |
| 4 | Destination | Users exist and licensed for OneDrive, no OneDrive yet | Docs |
| 5 | Destination | Upload the identity map | Docs |
| 6 | Source | Check compatibility, queue the moves, watch them | Docs |

```powershell
# 1. Both tenants: note each URL (it looks like https://<tenant>-my.sharepoint.com/)
Get-SPOCrossTenantHostUrl

# 2. Source first, then destination
Set-SPOCrossTenantRelationship -Scenario MnA -PartnerRole Target -PartnerCrossTenantHostUrl https://contoso-my.sharepoint.com/   # in the source
Set-SPOCrossTenantRelationship -Scenario MnA -PartnerRole Source -PartnerCrossTenantHostUrl https://fabrikam-my.sharepoint.com/  # in the destination

# 3. Both: carry on only at GoodToProceed
Verify-SPOCrossTenantRelationship -Scenario MnA -PartnerRole Target -PartnerCrossTenantHostUrl https://contoso-my.sharepoint.com/   # in the source
Verify-SPOCrossTenantRelationship -Scenario MnA -PartnerRole Source -PartnerCrossTenantHostUrl https://fabrikam-my.sharepoint.com/  # in the destination
```

There's no `Set-SPOCrossTenantHostUrl` in Microsoft's steps; `Get-` gives you the URL and that's all you need. **4.** The destination users must exist (the mail users from step 6 above) and be licensed for OneDrive, with no OneDrive created. See the licensing traps above.

**5. Identity map.** A CSV with **no header row**, six columns: `User, <source tenant ID>, <source UPN>, <destination UPN>, <destination email>, <user type>`. The source tenant ID is on the source's Entra ID > Overview (or Properties) page. Build it from the mapping:

```powershell
$srcTenant = '<source tenant ID>'
Import-Csv .\UserMapping.csv | ForEach-Object { "User,$srcTenant,$($_.Source),$($_.Destination),$($_.Destination),Member" } |
    Set-Content .\IdentityMap.csv -Encoding utf8
Add-SPOTenantIdentityMap -IdentityMapPath .\IdentityMap.csv   # in the destination
```

Check Microsoft's step 5 page for the user type values before you upload; the example there is an image. The `Source` column must be the source **UPN**, which isn't always the primary address: check against `03-Users.csv` from [01](../../01-SourceInventory/). Each upload replaces the last one, so always upload the full list. Close the file in Excel first, or the upload says it's in use.

```powershell
# 6. In the source
Get-SPOCrossTenantCompatibilityStatus -PartnerCrossTenantHostURL https://contoso-my.sharepoint.com/   # Compatible or Warning: go
Start-SPOCrossTenantUserContentMove -SourceUserPrincipalName alex@fabrikam.onmicrosoft.com `
    -TargetUserPrincipalName alex@contoso.com -TargetCrossTenantHostUrl https://contoso-my.sharepoint.com/
# optional: -PreferredMoveBeginDate / -PreferredMoveEndDate, in UTC

Get-SPOCrossTenantUserContentMoveState -PartnerCrossTenantHostURL https://contoso-my.sharepoint.com/
Get-SPOCrossTenantUserContentMoveState -PartnerCrossTenantHostURL https://contoso-my.sharepoint.com/ -SourceUserPrincipalName alex@fabrikam.onmicrosoft.com -Verbose
```

States run NotStarted, Scheduled, InProgress, Success (or Failed, Rescheduled). `Stop-SPOCrossTenantUserContentMove` cancels one that hasn't started.

**Limits from Microsoft's docs:** up to 4,000 queued at once; 5 TB or 1 million items per OneDrive (versions count as items); paths up to 400 characters including the new URL; an apostrophe in a user name or URL fails; source OneDrives on hold, or read-only, fail; no Customer Key in the source.

## Three ways to do it

| Route | How |
| --- | --- |
| Paid tool | [B. ShareGate](../B-ShareGate/) or similar: copies into ordinary mailboxes built by [02](../../02-TargetMailboxes/). Archives over EWS, incrementals, no add-on |
| Microsoft's own | This page. Commands above, or Microsoft's cross-tenant migration orchestrator, which wraps the same steps |
| By hand | The commands above, with the mail users stamped one at a time in PowerShell (`New-MailUser`, then `Set-MailUser -ExchangeGuid -ArchiveGuid -EmailAddresses @{Add='x500:...'}`). Fine for three people. Past that, one wrong GUID costs you an evening |

## Tips and traps

- **Trap:** license an unstamped destination user and Exchange makes an empty mailbox; the move then fails. Stamp, move, license.
- **Trap:** 02 and 04 build mailboxes and OneDrives. Both make the native move fail for that person. Run neither for people moving natively.
- **Trap:** OneDrive is moved, not copied. There's no incremental, and the source OneDrive is gone afterwards (a redirect is left). Plan it for after cutover, when sign-in names are final, and tell people.
- **Trap:** ExternalEmailAddress on its own overwrites the primary SMTP. The script always sets both together; do the same by hand.
- **Tip:** a mailbox that's been in the source for years carries x500 addresses from older systems. Copy all of them, not just the LegacyExchangeDN, or replies to old messages bounce. The script does.
- **Tip:** start the first batch early. Native moves run on Microsoft's schedule, and the first pass of a big archive can take days. Completion is quick; syncing isn't.
- **Tip:** keep the app secret out of scripts and history. `Read-Host -AsSecureString`, as above, and delete the secret from the app registration when you've finished.

## Microsoft documentation

- [Cross-tenant mailbox migration](https://learn.microsoft.com/microsoft-365/enterprise/cross-tenant-mailbox-migration)
- [Cross-tenant OneDrive migration](https://learn.microsoft.com/microsoft-365/enterprise/cross-tenant-onedrive-migration)
- [Cross-tenant SharePoint migration](https://learn.microsoft.com/microsoft-365/enterprise/cross-tenant-sharepoint-migration)
- [Cross-tenant orchestrated user data migration overview](https://learn.microsoft.com/microsoft-365/migration/migration-orchestrator-1-overview)
