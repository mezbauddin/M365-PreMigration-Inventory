# M365 Pre-Migration Inventory

**TL;DR:** `Invoke-SourceTenantInventory.ps1` is a read-only PowerShell script that inventories a Microsoft 365 source tenant before a tenant-to-tenant migration. One run gives you mailbox and archive sizes, mailbox permissions, users and licences, groups, distribution lists, mail contacts, domains, SharePoint sites and OneDrive storage as CSV files, plus a one-page summary. It changes nothing in the tenant.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.0.3 |
| Released | 4 November 2026 |
| Licence | MIT |

Part 1 of my [M365 tenant-to-tenant migration toolkit](../README.md). Next: [New-TargetMailboxes](../02-TargetMailboxes/) builds the destination from this inventory.

## Why

Migration tools such as ShareGate scan SharePoint and Teams well, but the numbers you need to plan mail and identity (mailbox sizes, archives, holds, Full Access and Send As, distribution list members, mail contacts) sit in Exchange and Entra. This script collects those in one pass so you can size the job before you copy anything.

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
.\Invoke-SourceTenantInventory.ps1 -SharePointAdminUrl https://contoso-admin.sharepoint.com
```

On large tenants, skip the permissions scan first and run it separately later:

```powershell
.\Invoke-SourceTenantInventory.ps1 -SharePointAdminUrl https://contoso-admin.sharepoint.com -SkipMailboxPermissions
```

## Output

Files land in `.\Inventory-<tenant>-<yyyyMMdd-HHmm>\`.

| File | Contents |
| --- | --- |
| 00-Summary.csv | Counts and totals: users, mailbox types, mail and archive GB, largest mailbox, holds, lists, contacts, SharePoint and OneDrive GB |
| 01-AcceptedDomains.csv | Domains that must move or be retired |
| 02-LicenceSkus.csv | Purchased and consumed licences |
| 03-Users.csv | Users, licences, enabled state, last sign-in |
| 04-Mailboxes.csv | Every mailbox with type, size, item count, archive size, holds, forwarding and aliases |
| 05-MailboxPermissions.csv | Full Access and Send As grants (not copied by most migration tools) |
| 06-M365Groups.csv | Microsoft 365 groups and their sites |
| 07-DistributionLists.csv | Lists, members and whether outside senders are allowed |
| 08-MailContacts.csv | External contacts in the directory |
| 09-TransportRules.csv | Mail flow rules to recreate |
| 10-EwsConfig.csv | EWS state and allow list (matters for archive and group mailbox copies until 1 April 2027) |
| 11-SitesAndOneDrive.csv | Every site and OneDrive with storage used. This figure includes version history and the recycle bin |

## Notes

- Sizes are decimal GB (1 GB = 1,000,000,000 bytes) for Exchange. SharePoint storage comes from `StorageUsageCurrent` in MB and is divided by 1,024.
- The script only reads. It doesn't change any setting, permission or object.
- Not yet tested against a live tenant in this release. Run it against a test tenant first if you can, and open an issue with anything that breaks.

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 1.0.3 | 2026-11-04 | Fixes "cannot call a method on a null-valued expression" in the mailbox step (`ExchangeGuid` is now requested explicitly). A mailbox that can't be read gives a warning and the run carries on |
| 1.0.2 | 2026-11-04 | Checks modules before connecting. Loads the SharePoint module through `-UseWindowsPowerShell` on PowerShell 7. If that module is missing, the run carries on and skips only SharePoint and OneDrive |
| 1.0.1 | 2026-11-04 | `-SharePointAdminUrl` accepts the bare host (`contoso-admin.sharepoint.com`) as well as the full https URL |
| 1.0.0 | 2026-11-04 | First release |
