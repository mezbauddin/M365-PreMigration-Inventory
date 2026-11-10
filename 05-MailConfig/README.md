# Copy-MailConfig

**TL;DR:** Migration tools copy mail, not the setup around it. This script exports who can open whose mailbox, who can send as whom, who can see or edit whose calendar (delegates included), forwarding, distribution lists and contacts from the source, then rebuilds them in the destination using your mapping file. It's safe to re-run, and everything supports `-WhatIf`. Transport rules are exported for you to rebuild by hand, and inbox rules and automatic replies are listed for you to check after cutover. Examples use fabrikam.com as the source and contoso.com as the destination.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.2.0 |
| Released | 10 November 2026 |
| Licence | MIT |

Part 5 of my [M365 tenant-to-tenant migration toolkit](../README.md). Run the export any time before cutover. Run the import once the destination mailboxes exist, built with [New-TargetMailboxes](../02-TargetMailboxes/).

## Why you need this

A shared mailbox only works if the right people can open it. A team address only works if the list exists, with the right members. Mailbox copy tools, whichever one you use, move the content and leave all of that behind.

## What happens without it

- **Shared mailboxes arrive, and nobody can open them.** Full Access and Send As don't come with the mail. People find out on day one, when they can't reach the team inbox.
- **The PA can't see the boss's diary.** Calendar permissions and delegates live on the calendar folder. Most copy tools leave them behind, so meeting requests stop reaching the delegate and shared calendars go blank.
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
| Calendar permissions | `Get-MailboxFolderPermission` on each mailbox's calendar, delegates' flags included. Inbox and Contacts too with `-IncludeFolders`. Default and Anonymous only when changed from a fresh mailbox's | `Add-MailboxFolderPermission`, or `Set-MailboxFolderPermission` when the person already has a different right. Found by folder type, so a calendar called Kalender or Agenda still works |
| Inbox rules and automatic replies | Report only: `CheckAfterCutover.csv`, per mailbox with rules (and how many forward or redirect) or an automatic reply on or scheduled | Not imported. Check after cutover |
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

# Inbox and Contacts permissions as well as Calendar
.\Copy-MailConfig.ps1 -Export -TenantId <source-tenant-guid> -IncludeFolders Inbox, Contacts

# Big tenant, first pass: skip the per-mailbox folder and rule reads
.\Copy-MailConfig.ps1 -Export -TenantId <source-tenant-guid> -SkipFolderPermissions -SkipMailboxChecks

# Folder permissions only, once the mailboxes exist in the destination
.\Copy-MailConfig.ps1 -Import ... -Include FolderPermissions -WhatIf

# Name the account, so an open session for the other tenant is never reused
.\Copy-MailConfig.ps1 -Export -TenantId <source-tenant-guid> -AdminUpn admin@fabrikam.onmicrosoft.com
.\Copy-MailConfig.ps1 -Import -TenantId <destination-tenant-guid> -AdminUpn admin@contoso.onmicrosoft.com -ExportFolder ... -MappingFile ... -AddressDomain contoso.com
```

**Tip: name the account.** With both tenants open on one PC, `Connect-ExchangeOnline` and `Connect-MgGraph` happily reuse whichever session is already there. Add `-AdminUpn admin@contoso.onmicrosoft.com` and the script disconnects any session for another account or tenant, signs in as the one you named, and prints `Signed in as X to tenant Y` before it changes anything. On the live cutover this would have saved a confusing error at the worst moment.

## Output

- **Export:** a `MailConfig-<domain>-<date>` folder of CSVs, plus `TransportRules.xml`. New in 1.2.0: `FolderPermissions.csv` (mailbox, folder, trustee, rights, delegate flags) and `CheckAfterCutover.csv` (inbox rules and automatic replies to check).
- **Import:** `MailConfig-Import-<domain>-<date>.csv`, with every action marked Created, Added, Set, Exists, Skipped, Problem or WhatIf, and why.

The CSVs contain names and addresses. Keep them out of source control and shared drives. The repo's `.gitignore` already excludes `*.csv`.

## Transport rules

Rebuild these by hand, on purpose:

- `Import-TransportRuleCollection` is retired in Exchange Online. On-premises, it replaced every existing rule in the organisation.
- Rules usually name source-only things: domains, groups, specific people, connectors.
- A wrong mail flow rule can block, redirect or copy mail for the whole tenant.

`TransportRules.xml` holds a `New-TransportRule` command for each rule. Go through them one by one, fix the addresses, run each with `-Enabled $false`, test it, then turn it on.

### After an acquisition, leaving them behind is often right

When one company has bought another, the old company's mail flow rules are usually history, not configuration. On the live migration the source had years of them: a disclaimer for a brand that no longer traded, routing to a system that had been switched off, rules naming people who had left. The destination also served other companies, so a broad rule could have touched their mail too.

What I'd do every time now:

1. **Export them as reference only.** `TransportRules.csv` and `TransportRules.xml` are your record of what existed. Keep them with the handover.
2. **Ask the business which ones still matter.** Usually it's one or two: a disclaimer, a block on a known bad sender, a copy to a compliance mailbox. Often it's none.
3. **Rebuild only those, disabled.** Scope them to the moved users or domain, not the whole tenant. Test with a real message, then enable.
4. **Write down what you left behind and why,** so nobody wonders in six months.

**Trap:** a rule that "adds a disclaimer to all outbound mail" in the source becomes "adds the old company's disclaimer to every company's mail" in a shared destination.

### Forwarding to outside addresses is the business owner's call

The export lists every mailbox that forwards outside the organisation. Some are legitimate (an accountant, a partner firm). Some are a leaver's mail going to a personal address, which is a data protection problem waiting to happen. The script won't recreate them unless you pass `-IncludeExternalForwarding`, and I wouldn't pass it without a written yes from the business owner for each one. Outbound spam policy in the destination may block automatic forwarding anyway, so test any you do recreate.

## Inbox rules and automatic replies: check, don't copy

These belong to the mailbox, so whether they arrive depends on what moved the mail:

| Copy route | Inbox rules | Automatic replies |
| --- | --- | --- |
| [Microsoft native](../copy-options/A-Microsoft-native/) | Moved. It's a real mailbox move | Moved |
| [ShareGate](../copy-options/B-ShareGate/) | Copied, apart from a few rule options (run before, run after, pin to top). Mail flow rules aren't | Copied, apart from the calendar-blocking options (block my calendar, decline new invitations, decline and cancel my meetings) |
| [No budget](../copy-options/C-No-budget/) and most other tools | Usually not. Check your tool's limits page | Usually not |

Either way, a rule that forwards to `sam@fabrikam.com` or moves mail into a folder that didn't arrive breaks quietly. `CheckAfterCutover.csv` lists every mailbox with rules or an automatic reply, with a plain-English line on what to check. Go through it on T+1 with the people concerned.

## Three ways to do it

| Route | How |
| --- | --- |
| Paid tool | Mailbox copy tools move the content. ShareGate doesn't keep folder permissions, and few tools rebuild Full Access, Send As, calendar delegates, lists or forwarding in the destination. Check what yours does, then use this for the rest |
| Microsoft's own | Native cross-tenant mailbox migration moves the mailbox with its own folder permissions, but not Full Access or Send As on other mailboxes, distribution lists or contacts. Exchange admin centre for each one by hand |
| By hand | Exchange admin centre: for each shared mailbox, Delegation, add Full Access, Send As, Send on Behalf. For each list, recreate and add members. Calendars: each person shares from Outlook again, or `Add-MailboxFolderPermission` one at a time. Fine for ten objects; error-prone for a hundred |

## Lessons learnt

- **Map before you import.** Everything hinges on the mapping file. A permission is only rebuilt if both the mailbox and the person are in it. The import report lists every skip, so check it.
- **Deleted people show up as trustees.** Old permissions often point at accounts that no longer exist. The export marks them `unresolved:` and the import skips them. That's usually the right call.
- **Alias clashes are normal.** A destination that's been running for years will already have `info`, `sales` or `accounts`. The script won't take over an existing list. Use `-AliasPrefix`, and sort out the real address at cutover.
- **Contacts may already exist.** If the destination already has a recipient with a contact's external address, the script reuses it rather than creating a duplicate.
- **External forwarding is a decision, not a default.** It might be a leaver's mail going to a personal address. Check with the business before switching it back on, and remember outbound spam policy may block automatic forwarding anyway.
- **Tip: calendar permissions need the mailbox to be there.** Run `-Include FolderPermissions` after 02 has built the mailboxes. A staged shared mailbox has a calendar, so you don't have to wait for cutover.
- **Trap: delegates on a staged mailbox.** Delegate flags (`Delegate`, `CanViewPrivateItems`) only apply to Editor on a calendar. If a delegate stops getting meeting requests after cutover, check the person is a user mailbox by then and re-run `-Include FolderPermissions`.
- **Trap: calendars shared outside the organisation** are listed as `external:` and skipped. The owner re-shares them from the destination; the old sharing link points at the source.
- **Trap: rules pointing at old addresses.** Native moves carry inbox rules, but a rule that forwards to someone by their old address only works if that address is on the new mailbox. `CheckAfterCutover.csv` tells you whose rules to look at.
- **Full Access automapping is on.** Shared mailboxes appear in Outlook on their own. If someone has Full Access to dozens of mailboxes, consider turning it off for them.

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 1.2.0 | 2026-11-10 | Calendar permissions and delegates (Inbox and Contacts with `-IncludeFolders`) exported and rebuilt through the mapping, by folder type. `CheckAfterCutover.csv`: inbox rules and automatic replies per mailbox, report only. `-SkipFolderPermissions`, `-SkipMailboxChecks`, and `FolderPermissions` in `-Include` |
| 1.1.0 | 2026-11-10 | `-AdminUpn`: disconnects a session for any other account or tenant, signs in as the one named, prints who and where. Guidance on leaving transport rules behind after an acquisition, and on external forwarding |
| 1.0.0 | 2026-11-09 | First release |
