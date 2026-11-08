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

More to come as the migration I'm running moves on: OneDrive and SharePoint mapping, and the cutover steps.

EWS in Exchange Online is retired on 1 April 2027. Scripts and tools that copy archives over EWS stop working then.

## How they fit together

1. Run **Invoke-SourceTenantInventory** against the source. It writes `Inventory-<tenant>-<date>\04-Mailboxes.csv` and friends.
2. Run **New-TargetMailboxes -Plan** in the same folder. It finds the inventory and writes `TargetBuildPlan.csv`. Review it.
3. Run **New-TargetMailboxes -Apply -StageOnly** against the destination, then start your mail copy.
4. Before copying archives, run **Test-EwsArchiveReadiness** with your migration tool's app ID and both tenant IDs. It checks and fixes both tenants in one run.

## A word on the output

The CSVs these scripts write hold tenant data: names, addresses, permissions and, for script 2, temporary passwords. The `.gitignore` keeps them out of commits. Keep them out of email and shared drives too.

## Older links

This repo used to be called M365-PreMigration-Inventory. GitHub redirects the old URL here. The inventory script now lives in [01-SourceInventory](01-SourceInventory/).
