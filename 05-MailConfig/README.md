# Copy-MailConfig

**TL;DR:** Migration tools copy mail, not the setup around it. This script exports who can open whose mailbox, who can send as whom, forwarding, distribution lists and contacts from the source, then rebuilds them in the destination using your mapping file. It's safe to re-run, and everything supports `-WhatIf`. Transport rules are exported for you to rebuild by hand.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.0.0 |
| Released | 9 November 2026 |
| Licence | MIT |

Part 5 of my [M365 tenant-to-tenant migration toolkit](../README.md). Run the export any time before cutover. Run the import once the destination mailboxes exist, built with [New-TargetMailboxes](../02-TargetMailboxes/).

## Why you need this

A shared mailbox only works if the right people can open it. A team address only works if the list exists, with the right members. Mailbox copy tools, whichever one you use, move the content and leave all of that behind.

## What happens without it

- **Shared mailboxes arrive, and nobody can open them.** Full Access and Send As don't come with the mail. People find out on day one, when they can't reach the team inbox.
- **Replies go out from the wrong address.** Without Send As or Send on Behalf, people reply to customers from their own mailbox.
- **Mail to team addresses bounces,** or lands nowhere, because the distribution lists don't exist in the destination.
- **Forwarding silently stops.** Mail meant for a manager, an accounts system or an archive address goes nowhere, and nobody notices for weeks.
- **External contacts vanish from the address book,** and any list that included them loses those members.

## What it does

| Part | Export | Import |
| --- | --- | --- |
| Full Access | Per mailbox, explicit grants only (not inherited, not system accounts) | `Add-MailboxPermission` with automapping |
| Send As | Per mailbox | `Add-RecipientPermission` |
| Send on Behalf | Per mailbox | `Set-Mailbox -GrantSendOnBehalfTo @{Add=...}` |
| Forwarding | Internal and external, with 'keep a copy' | Internal recreated. External only with `-IncludeExternalForwarding` |
| Distribution lists | Settings, owners, all addresses, members | Created on `-AddressDomain`, owners and members mapped |
| Dynamic lists | Name, address and filter | Not imported. Filters rarely survive a move unchanged |
| Mail contacts | Name, external address, hidden flag | Created, or an existing recipient with the same address reused |
| Transport rules | Summary CSV plus the full XML | Not imported. See below |

Every object the import creates is tagged with `CustomAttribute15 = T2T:<source address>`. That's how re-runs know what's already done, and how the cutover step can find which object gets which source address.

## Requirements

- `ExchangeOnlineManagement` 3.x, as Exchange Administrator in each tenant
- A mapping CSV with `Source` and `Destination` columns: source mailbox address to destination sign-in name or address. The same file works for SharePoint and Teams user mapping
- An accepted domain in the destination for new list addresses, e.g. a staging domain. Source domains get added at cutover

## Usage

```powershell
# 1. Source tenant: export. Changes nothing.
.\Copy-MailConfig.ps1 -Export -TenantId <source-tenant-guid>

# 2. Destination tenant: preview, then import
Disconnect-ExchangeOnline -Confirm:$false
.\Copy-MailConfig.ps1 -Import -TenantId <destination-tenant-guid> -ExportFolder .\MailConfig-fabrikam.com-<date> `
    -MappingFile .\UserMapping.csv -AddressDomain contoso.com -WhatIf
.\Copy-MailConfig.ps1 -Import -TenantId <destination-tenant-guid> -ExportFolder .\MailConfig-fabrikam.com-<date> `
    -MappingFile .\UserMapping.csv -AddressDomain contoso.com

# Destination already has lists called 'info' or 'sales'?
.\Copy-MailConfig.ps1 -Import ... -AliasPrefix fab-

# Only some parts
.\Copy-MailConfig.ps1 -Import ... -Include Permissions, Forwarding
```

## Output

- **Export:** a `MailConfig-<domain>-<date>` folder of CSVs, plus `TransportRules.xml`.
- **Import:** `MailConfig-Import-<domain>-<date>.csv`, with every action marked Created, Added, Set, Exists, Skipped, Problem or WhatIf, and why.

The CSVs contain names and addresses. Keep them out of source control and shared drives. The repo's `.gitignore` already excludes `*.csv`.

## Transport rules

Rebuild these by hand, on purpose:

- `Import-TransportRuleCollection` is retired in Exchange Online. On-premises, it replaced every existing rule in the organisation.
- Rules usually name source-only things: domains, groups, specific people, connectors.
- A wrong mail flow rule can block, redirect or copy mail for the whole tenant.

`TransportRules.xml` holds a `New-TransportRule` command for each rule. Go through them one by one, fix the addresses, run each with `-Enabled $false`, test it, then turn it on.

## Lessons learnt

- **Map before you import.** Everything hinges on the mapping file. A permission is only rebuilt if both the mailbox and the person are in it. The import report lists every skip, so check it.
- **Deleted people show up as trustees.** Old permissions often point at accounts that no longer exist. The export marks them `unresolved:` and the import skips them. That's usually the right call.
- **Alias clashes are normal.** A destination that's been running for years will already have `info`, `sales` or `accounts`. The script won't take over an existing list. Use `-AliasPrefix`, and sort out the real address at cutover.
- **Contacts may already exist.** If the destination already has a recipient with a contact's external address, the script reuses it rather than creating a duplicate.
- **External forwarding is a decision, not a default.** It might be a leaver's mail going to a personal address. Check with the business before switching it back on, and remember outbound spam policy may block automatic forwarding anyway.
- **Full Access automapping is on.** Shared mailboxes appear in Outlook on their own. If someone has Full Access to dozens of mailboxes, consider turning it off for them.

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 1.0.0 | 2026-11-09 | First release |
