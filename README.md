# M365 Tenant-to-Tenant Migration Toolkit

**TL;DR:** PowerShell scripts I use on Microsoft 365 tenant-to-tenant migrations, in the order I run them. Inventory the source first, then build the destination so your migration tool has somewhere to copy to. Each script has its own README with a TL;DR, requirements and the gotchas I hit.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Licence | MIT |

## The scripts

| # | Script | What it does | Changes anything? |
| --- | --- | --- | --- |
| 1 | [Invoke-SourceTenantInventory](01-SourceInventory/) | Mailbox and archive sizes, permissions, users and licences, groups, lists, contacts, domains, SharePoint and OneDrive, as CSVs | No, read-only |
| 2 | [New-TargetMailboxes](02-TargetMailboxes/) | Builds the destination mailboxes from the inventory. Stage as shared, copy, then license and convert at cutover | Yes, destination only. Supports `-WhatIf` |
| 3 | [Test-EwsArchiveReadiness](03-EwsArchiveReadiness/) | Signs in to source and destination in turn and finds what stops archives copying over EWS: tenant switch, app ID allow list, old user-agent rules, per-mailbox blocks, consent, archive readiness. Without it, main mail copies and every archive fails | Only with `-Fix`. Supports `-WhatIf` |
| 4 | [Request-TargetOneDrives](04-TargetOneDrives/) | Creates the destination OneDrives for everyone in the plan, in batches of 200, and reports each URL. Without it, the OneDrive copy fails for anyone who hasn't signed in yet | Yes, creates OneDrives. Supports `-WhatIf` |
| 5 | [Copy-MailConfig](05-MailConfig/) | Exports mailbox permissions, forwarding, distribution lists and contacts from the source and rebuilds them in the destination. Transport rules exported for review. Without it, shared mailboxes arrive that nobody can open | Export: no. Import: yes, destination only. Supports `-WhatIf` |

More to come as the migration I'm running moves on: the domain cutover, validation, device rejoin, source decommissioning and a handover pack.

EWS in Exchange Online is retired on 1 April 2027. Scripts and tools that copy archives over EWS stop working then.

## How they fit together

1. Run **Invoke-SourceTenantInventory** against the source. It writes `Inventory-<tenant>-<date>\04-Mailboxes.csv` and friends.
2. Run **New-TargetMailboxes -Plan** in the same folder. It finds the inventory and writes `TargetBuildPlan.csv`. Review it.
3. Run **New-TargetMailboxes -Apply -StageOnly** against the destination.
4. If your tool copies In-Place Archives, run **Test-EwsArchiveReadiness** with its app ID and both tenant IDs **before** you start the mail copy. Then start the copy.
5. Once people are licensed, run **Request-TargetOneDrives** so their OneDrives are ready. Copy them after cutover.
6. Run **Copy-MailConfig -Export** against the source, then **-Import** against the destination, so shared mailboxes, lists and forwarding work on day one.

## Two tenants, one PC

You'll be signed in to two tenants from the same machine for weeks. Windows quietly reuses your last sign-in, so `Connect-ExchangeOnline` or `Connect-MgGraph` can land you in the wrong tenant without asking. Scripts in this toolkit that change things check the tenant ID before they do. When you run commands by hand, name the account (`-UserPrincipalName`) or the tenant (`-TenantId`), and check `Get-ConnectionInformation` before anything destructive.

## A word on the output

The CSVs these scripts write hold tenant data: names, addresses, permissions and, for script 2, temporary passwords. The `.gitignore` keeps them out of commits. Keep them out of email and shared drives too.

## Older links

This repo used to be called M365-PreMigration-Inventory. GitHub redirects the old URL here. The inventory script now lives in [01-SourceInventory](01-SourceInventory/).
