# Option A: Microsoft native cross-tenant migration

**TL;DR:** Microsoft can move mailboxes (with archives) and OneDrive between tenants itself, server to server, keeping items, folders and permissions. It needs a paid per-user add-on, a fair bit of setup across both tenants, and destination users created as mail users with the right identity attributes, not as mailboxes. It doesn't move shared Teams content, and its SharePoint add-on was Enterprise Agreement only at the time of writing. This page is the plan. Scripts follow once each step has been run in my lab tenants.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Status | Being lab-tested between two tenants. Scripts land here when they've been run end to end |

Part of the [copy options](../README.md) in my [M365 tenant-to-tenant migration toolkit](../../README.md).

## What it covers

| Workload | Native route | Licence |
| --- | --- | --- |
| Mailboxes and In-Place Archives | Cross-tenant mailbox migration (migration batches run from the destination) | Cross-tenant user data migration add-on, per user, one-time |
| OneDrive | Cross-tenant OneDrive migration (SharePoint Online PowerShell) | Same add-on |
| SharePoint sites | Cross-tenant SharePoint migration | Cross-tenant shared data migration, per 100 GB. Enterprise Agreement only at the time of writing |
| Teams channels, chat in channels, Planner, Power Platform | Not covered | n/a |

Check Microsoft's current pricing and availability before you quote. Both have changed during 2026.

## How it differs from the rest of this toolkit

The biggest change is **how the destination accounts are built**.

- **[02 Target mailboxes](../../02-TargetMailboxes/) creates mailboxes.** That's right for ShareGate and the no-budget route.
- **Native mailbox migration needs mail users.** These are destination accounts with no mailbox yet, stamped with the source mailbox's ExchangeGuid, ArchiveGuid, LegacyExchangeDN (as an X500) and addresses. The migration then moves the mailbox into that object.
- **Licence too early and you get a fresh, empty mailbox.** The migration for that person then fails, and you're unpicking it by hand. So with native: create mail users, stamp them, migrate, then license.

The other parts of the spine are unchanged: [05 Mail config](../../05-MailConfig/), [07 Domain cutover](../../07-DomainCutover/). [03 EWS readiness](../../03-EwsArchiveReadiness/) isn't needed, because native migration doesn't use EWS.

## The plan, in order

1. **Licences.** Buy the cross-tenant user data migration add-on for everyone moving. Assign it in either tenant.
2. **Destination: migration app.** An Entra app registration with the `Mailbox.Migration` permission and a secret or certificate. Consent in the destination.
3. **Source: consent and organisation relationship.** Consent the app from the source. Create an organisation relationship allowing the move from the source, scoped to a mail-enabled security group of the mailboxes moving.
4. **Destination: organisation relationship and migration endpoint.** The matching relationship, and a migration endpoint pointing at the source with the app.
5. **Destination: mail users.** Create one per mailbox, stamped with the source ExchangeGuid, ArchiveGuid, X500 of the LegacyExchangeDN, and the addresses. This is the step to script, and the one people get wrong.
6. **Test batch.** One or two mailboxes. Check items, archive, calendar, permissions.
7. **Batches.** Move in waves. Complete each just before cutover.
8. **OneDrive.** Set up the trust between tenants in SharePoint Online PowerShell, upload the identity map, then queue the OneDrive moves.
9. **Cutover.** [07 Domain cutover](../../07-DomainCutover/), as for the other options.

## Scripts planned for this folder

| Script | Purpose |
| --- | --- |
| `New-CrossTenantMigrationApp.ps1` | Destination app registration, permission, secret, consent link for the source |
| `Set-CrossTenantRelationship.ps1` | Organisation relationships in both tenants, scoped to a security group |
| `New-TargetMailUsers.ps1` | Mail users in the destination, stamped from the source (ExchangeGuid, ArchiveGuid, X500, addresses) |
| `Start-CrossTenantMailboxBatch.ps1` | Endpoint, batch, status and completion |
| `Start-CrossTenantOneDriveMove.ps1` | Trust, identity map, OneDrive moves and status |

## Microsoft documentation

- [Cross-tenant mailbox migration](https://learn.microsoft.com/microsoft-365/enterprise/cross-tenant-mailbox-migration)
- [Cross-tenant OneDrive migration](https://learn.microsoft.com/microsoft-365/enterprise/cross-tenant-onedrive-migration)
- [Cross-tenant SharePoint migration](https://learn.microsoft.com/microsoft-365/enterprise/cross-tenant-sharepoint-migration)
- [Cross-tenant orchestrated user data migration overview](https://learn.microsoft.com/microsoft-365/migration/migration-orchestrator-1-overview)
