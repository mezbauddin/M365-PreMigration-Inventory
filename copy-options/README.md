# Copy options: three ways to move the data

**TL;DR:** Every tenant-to-tenant migration has the same spine (inventory, build the destination, rebuild the config, cut over the domain, validate, decommission). The only part that really changes is **how the data gets copied**. This folder covers three honest options: Microsoft's own cross-tenant migration, ShareGate, and a no-budget manual route. Pick by what you're moving and what you can spend, not by habit.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Status | B used in production on a live migration. A being lab-tested. C written up from standard practice |

Part of my [M365 tenant-to-tenant migration toolkit](../README.md).

## The short version

| | [A. Microsoft native](A-Microsoft-native/) | [B. ShareGate](B-ShareGate/) | [C. No budget](C-No-budget/) |
| --- | --- | --- | --- |
| Mailboxes and archives | Yes, cross-tenant mailbox migration | Yes | PST export and import. Slow, manual |
| OneDrive | Yes, cross-tenant OneDrive migration | Yes | Download and re-upload, or sync client |
| SharePoint sites | Yes, but the add-on is Enterprise Agreement only | Yes | Download and re-upload. Loses versions and metadata |
| Teams channels and chat | Not shared Teams content | Teams, channels and files | Files only |
| Cost | Add-on per user, one-time. SharePoint priced per 100 GB | Subscription. Check current pricing | Your time |
| Permissions, versions, authors | Kept | Kept, through user mapping | Mostly lost |
| Incremental passes | Yes | Yes | No |
| Setup effort | High: app registration, trust between tenants, identity mapping | Low: sign in to both tenants | Low setup, high labour |
| Best for | Larger mail and OneDrive moves, EA customers | Small to mid businesses, anything with SharePoint and Teams | A handful of users with little shared data |

Microsoft's add-on prices change, so check the current list before quoting. At the time of writing, cross-tenant user data migration was a per-user one-time add-on covering mailbox and OneDrive, and shared data migration for SharePoint was sold per 100 GB to Enterprise Agreement customers.

## How to choose

1. **Is there much SharePoint or Teams content?** If yes, and you're not on an Enterprise Agreement, the native route has nothing for it. That points to B, or to C if it's small.
2. **How many people?** Native costs per user. Under about 50 people, the native setup effort often outweighs the licence saving.
3. **Do versions, authors and permissions matter?** If yes, rule out C for anything beyond a few users.
4. **Can you mix?** Yes. Native for mail and OneDrive with ShareGate for SharePoint is a common split. The shared spine works the same either way.

## What stays the same whichever you pick

These parts of the toolkit don't care how the data moves:

- [01 Source inventory](../01-SourceInventory/)
- [02 Target mailboxes](../02-TargetMailboxes/). For the native route, read the identity mapping note in A first, because native migration needs mail users, not mailboxes
- [04 Target OneDrives](../04-TargetOneDrives/)
- [05 Mail config](../05-MailConfig/)
- [07 Domain cutover](../07-DomainCutover/)

[03 EWS archive readiness](../03-EwsArchiveReadiness/) only applies if your tool copies archives over EWS. ShareGate does. Microsoft's native mailbox migration doesn't.

## Disclosure

I'm not sponsored by ShareGate or anyone else. I used ShareGate on the migration this toolkit came from because it fit the job, and I've written down exactly what it did well and where it tripped me up.
