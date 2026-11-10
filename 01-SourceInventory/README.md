# Invoke-SourceTenantInventory

**TL;DR:** `Invoke-SourceTenantInventory.ps1` is a read-only PowerShell script that inventories a Microsoft 365 source tenant before a tenant-to-tenant migration. One run gives you mailbox and archive sizes, mailbox permissions, users and licences, groups, distribution lists, mail contacts, domains, SharePoint sites and OneDrive storage as CSV files, plus a one-page summary. Since 1.1.0 it also lists public folders, inactive mailboxes, rooms and equipment with their booking rules, and the ExchangeGuid, ArchiveGuid and LegacyExchangeDN that a native cross-tenant move needs. It changes nothing in the tenant. Examples use fabrikam.com as the source, moving into the contoso tenant.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.1.0 |
| Released | 10 November 2026 |
| Licence | MIT |

Part 1 of my [M365 tenant-to-tenant migration toolkit](../README.md). Next: [New-TargetMailboxes](../02-TargetMailboxes/) builds the destination from this inventory.

## Why

Migration tools scan SharePoint and Teams well, but the numbers you need to plan mail and identity (mailbox sizes, archives, holds, Full Access and Send As, distribution list members, mail contacts) sit in Exchange and Entra. This script collects those in one pass so you can size the job before you copy anything.

## What happens without it

- **The copy runs for days longer than planned.** Archives are often most of the data. Without sizes you can't tell the business whether a weekend is enough.
- **Shared mailboxes arrive that nobody can open.** Full Access and Send As aren't copied by most tools. If you never listed them, you can't rebuild them.
- **Rooms stop taking bookings.** Room and equipment mailboxes, their capacity, delegates and booking rules are easy to forget until the first meeting request bounces.
- **Public folders get found on cutover night.** Nothing moves them by default, and the native route doesn't move them at all.
- **Held data gets deleted.** Inactive mailboxes and litigation holds don't show in the normal mailbox list. Close the source without knowing about them and you've broken a hold.
- **The native route stalls.** A cross-tenant move won't start unless the target mail user carries the source ExchangeGuid, ArchiveGuid and LegacyExchangeDN (as an X500). Collect them now, not when the move fails.

## Requirements

- Global Reader or Global Administrator in the source tenant
- PowerShell modules:
  - `ExchangeOnlineManagement` 3.x
  - `Microsoft.Graph.Users` and `Microsoft.Graph.Identity.DirectoryManagement`
  - `Microsoft.Online.SharePoint.PowerShell` (run from Windows PowerShell 5.1, or PowerShell 7 with `Import-Module -UseWindowsPowerShell`)
- Sign-in activity is optional. It needs Entra ID P1 and `AuditLog.Read.All`. Without it the script carries on and leaves `LastSignIn` blank.

```powershell
Install-Module ExchangeOnlineManagement, Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
```

## Usage

```powershell
.\Invoke-SourceTenantInventory.ps1 -SharePointAdminUrl https://fabrikam-admin.sharepoint.com
```

On large tenants, skip the permissions scan first and run it separately later:

```powershell
.\Invoke-SourceTenantInventory.ps1 -SharePointAdminUrl https://fabrikam-admin.sharepoint.com -SkipMailboxPermissions
```

## Output

Files land in `.\Inventory-<tenant>-<yyyyMMdd-HHmm>\`.

| File | Contents |
| --- | --- |
| 00-Summary.csv | Counts and totals: users, mailbox types, mail and archive GB, largest mailbox, holds, lists, contacts, SharePoint and OneDrive GB |
| 01-AcceptedDomains.csv | Domains that must move or be retired |
| 02-LicenceSkus.csv | Purchased and consumed licences |
| 03-Users.csv | Users, licences, enabled state, last sign-in |
| 04-Mailboxes.csv | Every mailbox with type, size, item count, archive size, holds, forwarding and aliases. From 1.1.0, also `ExchangeGuid`, `ArchiveGuid`, `LegacyExchangeDN` and `X500Addresses` at the end, for the [native route](../copy-options/A-Microsoft-native/) |
| 05-MailboxPermissions.csv | Full Access and Send As grants (not copied by most migration tools) |
| 06-M365Groups.csv | Microsoft 365 groups and their sites |
| 07-DistributionLists.csv | Lists, members and whether outside senders are allowed |
| 08-MailContacts.csv | External contacts in the directory |
| 09-TransportRules.csv | Mail flow rules to recreate |
| 10-EwsConfig.csv | EWS state and allow list (matters for archive and group mailbox copies until 1 April 2027) |
| 11-SitesAndOneDrive.csv | Every site and OneDrive with storage used. This figure includes version history and the recycle bin |
| 12-PublicFolders.csv | Every public folder, with path, type, mail-enabled flag and content mailbox. Only written if the tenant has public folder mailboxes |
| 13-MailPublicFolders.csv | Mail-enabled public folders and their addresses. These addresses must land somewhere at cutover |
| 14-InactiveMailboxes.csv | Mailboxes of deleted users kept by a hold. They don't move, and they decide what you can delete at [decommission](../10-Decommission/) |
| 15-ResourceMailboxes.csv | Rooms and equipment with capacity and booking rules: auto-accept, booking window, conflicts, delegates, who can book. Read by [New-TargetResourceMailboxes](../02-TargetMailboxes/) |

`00-Summary.csv` also gains `PublicFolders`, `MailPublicFolders` and `InactiveMailboxes` at the end. Existing files and columns are unchanged, so anything that read 1.0.x output still works.

## Three ways to do it

| Route | How |
| --- | --- |
| Paid tool | ShareGate's pre-migration reports and most migration suites scan SharePoint, Teams and OneDrive well and give mailbox sizes. Few list Full Access, Send As, room booking rules, inactive mailboxes or the GUIDs a native move needs. Use the tool's scan for content, this for mail and identity |
| Microsoft's own | Microsoft 365 admin centre: Reports, Usage (mailbox and OneDrive sizes), Active users export, Licences. Exchange admin centre for each shared mailbox's delegation and each room's booking settings. Purview for holds. All free, all one screen at a time |
| By hand | `Get-EXOMailbox`, `Get-EXOMailboxStatistics`, `Get-EXOMailboxPermission`, `Get-CalendarProcessing` and `Get-SPOSite` one by one, pasted into a spreadsheet. Fine for ten mailboxes. This script is those commands in one pass |

## Tips and traps

- **Tip:** run it twice. Once at the start to plan, and again around T-3 so the item counts 09 validates against are recent.
- **Tip:** open `15-ResourceMailboxes.csv` with whoever manages the rooms. Booking rules set years ago often surprise people, and it's the moment to fix them.
- **Trap:** `ResourceDelegates` or `BookInPolicy` showing `unresolved:` means the person has left. The rebuild skips them. Decide who should approve bookings now.
- **Trap:** storage in `11-SitesAndOneDrive.csv` includes version history and the recycle bin. The copy is usually smaller. Don't quote it to the business as "data to move".
- **Trap:** `14-InactiveMailboxes.csv` with any rows means a hold has kept data for people who've gone. Don't close the source until whoever owns that hold signs it off.
- **Trap:** a public folder count above zero needs its own plan. No copy option in this toolkit moves public folders. Usually they become a shared mailbox or a Microsoft 365 group, copied with a tool that supports public folders.
- **Trap:** on PowerShell 7 the SharePoint module loads through Windows PowerShell. If it fails, the run carries on without SharePoint. Read the warning before you trust a zero.

## Notes

- Sizes are decimal GB (1 GB = 1,000,000,000 bytes) for Exchange. SharePoint storage comes from `StorageUsageCurrent` in MB and is divided by 1,024.
- The script only reads. It doesn't change any setting, permission or object.
- Run against a live source tenant of about 40 mailboxes as part of a real migration. Run it against a test tenant first if you can, and open an issue with anything that breaks.

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 1.1.0 | 2026-11-10 | Adds public folders, mail-enabled public folders, inactive mailboxes, and rooms and equipment with capacity and calendar processing (12 to 15). `04-Mailboxes.csv` gains `ExchangeGuid`, `ArchiveGuid`, `LegacyExchangeDN` and `X500Addresses`; the summary gains three counts. Existing outputs unchanged. Examples now use fabrikam.com as the source |
| 1.0.4 | 2026-11-09 | `Aliases` lists secondary addresses only. The primary address was being listed as an alias too |
| 1.0.3 | 2026-11-04 | Fixes "cannot call a method on a null-valued expression" in the mailbox step (`ExchangeGuid` is now requested explicitly). A mailbox that can't be read gives a warning and the run carries on |
| 1.0.2 | 2026-11-04 | Checks modules before connecting. Loads the SharePoint module through `-UseWindowsPowerShell` on PowerShell 7. If that module is missing, the run carries on and skips only SharePoint and OneDrive |
| 1.0.1 | 2026-11-04 | `-SharePointAdminUrl` accepts the bare host (`fabrikam-admin.sharepoint.com`) as well as the full https URL |
| 1.0.0 | 2026-11-04 | First release |
