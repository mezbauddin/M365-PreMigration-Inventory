# Copy options: three ways to move the data

**TL;DR:** Every tenant-to-tenant migration has the same spine (inventory, build the destination, rebuild the config, cut over the domain, validate, decommission). The only part that really changes is **how the data gets copied**. This folder covers three honest options: Microsoft's own cross-tenant migration, ShareGate, and a no-budget manual route. Pick by what you're moving and what you can spend, not by habit.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 2.0.0 |
| Status | B used in production on a live migration. A is a full command list from Microsoft's documentation, not yet lab-verified. C written up from Microsoft's documentation and standard practice |
| Licence | MIT |

Part of my [M365 tenant-to-tenant migration toolkit](../README.md).

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 2.0.0 | 2026-11-10 | A gets a pasteable command list and `New-TargetMailUsers.ps1`. C gets the full PST path and `New-PstImportMapping.ps1`. Adds the scripts table, what happens without a choice, and version tables on every option |
| 1.1.0 | 2026-11-10 | `New-DriveCopyList.ps1` 1.0.0, shared by B and C |
| 1.0.0 | 2026-11-09 | First version: three options and a comparison |

## Scripts in this folder

| Script | Version | Option | Changes anything? |
| --- | --- | --- | --- |
| [`New-DriveCopyList.ps1`](New-DriveCopyList.ps1) | 1.0.0 | B and C | Only with `-CreateFolders` (destination folders). `-WhatIf` |
| [`A-Microsoft-native/New-TargetMailUsers.ps1`](A-Microsoft-native/New-TargetMailUsers.ps1) | 1.0.0 | A | Destination: creates and stamps mail users. `-WhatIf`. Never licenses |
| [`B-ShareGate/Copy-DrivesWithShareGate.ps1`](B-ShareGate/Copy-DrivesWithShareGate.ps1) | 1.1.0 | B | Destination: copies OneDrives. `-WhatIf` pre-check |
| [`C-No-budget/New-PstImportMapping.ps1`](C-No-budget/New-PstImportMapping.ps1) | 1.0.0 | C | No. Reads local PSTs, writes CSVs |
| [`C-No-budget/Copy-DriveViaGraph.ps1`](C-No-budget/Copy-DriveViaGraph.ps1) | 1.1.0 | C | Destination on `-Upload`. `-WhatIf` |

## Why choose deliberately

The copy method decides how the destination has to be built. ShareGate and the no-budget route copy into ordinary mailboxes, built by [02](../02-TargetMailboxes/). Microsoft's native route moves into stamped mail users and fails if a mailbox is already there. Pick late, and you may have to tear down what you built.

## What happens without a choice

- **The wrong destination objects.** Run 02 and 04 for everyone, then decide on native, and every person with a mailbox or OneDrive already in the destination fails their native move until you remove it.
- **EWS surprises.** Choose a tool that copies archives over EWS without knowing it, and archives fail until [03](../03-EwsArchiveReadiness/) is run in both tenants. After 1 April 2027 they won't work at all.
- **Budget arguments at the worst moment.** The native add-on and tool subscriptions take time to buy. Discovering you need one the week of cutover moves the date.

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

**Tip:** whatever you pick, decide before you run 02. **Trap:** if you mix, decide per person, not per workload, for mail: a person is either a native move (mail user) or a tool copy (mailbox), never both.

## What stays the same whichever you pick

These parts of the toolkit don't care how the data moves:

- [01 Source inventory](../01-SourceInventory/)
- [02 Target mailboxes](../02-TargetMailboxes/), except for anyone moving by the native route: they need stamped mail users from [A's `New-TargetMailUsers.ps1`](A-Microsoft-native/New-TargetMailUsers.ps1), not mailboxes
- [04 Target OneDrives](../04-TargetOneDrives/), again except for native OneDrive moves, which fail if the destination OneDrive already exists
- [05 Mail config](../05-MailConfig/)
- [06 Groups, Teams and sites](../06-GroupsTeamsSites/)
- [07 Domain cutover](../07-DomainCutover/)
- [09 Validation](../09-Validation/)
- [New-DriveCopyList.ps1](New-DriveCopyList.ps1), below, which builds the OneDrive copy list for B and C

[03 EWS archive readiness](../03-EwsArchiveReadiness/) only applies if your tool copies archives over EWS. ShareGate does. Microsoft's native mailbox migration doesn't.

## Building the OneDrive copy list: `New-DriveCopyList.ps1`

**TL;DR:** One script, shared by B and C, that writes the CSV both copy scripts read: every leaver's OneDrive into its own folder in an archive library, and current staff into their new OneDrive. It can create the folders too. Archived OneDrives are marked Skip with the reason, so a copy run doesn't fail on them.

| | |
| --- | --- |
| Script | [`New-DriveCopyList.ps1`](New-DriveCopyList.ps1) |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.0.0 |
| Changes anything? | No. Only with `-CreateFolders` (creates folders in the destination library), which supports `-WhatIf` |
| Used by | [B. Copy-DrivesWithShareGate.ps1](B-ShareGate/Copy-DrivesWithShareGate.ps1) and [C. Copy-DriveViaGraph.ps1](C-No-budget/Copy-DriveViaGraph.ps1) `-CsvPath` |

### Why you need it

Former staff usually don't get a OneDrive of their own in the destination: nobody licenses a leaver. The usual answer is one archive library on a SharePoint site, with a folder per person. On the live migration I built that CSV by hand (`Source, DestinationSite, DestinationLibrary, DestinationFolder, SourceLibrary`), worked out who had left by cross-checking licences and blocked accounts, and created each folder by clicking. It took an evening, and I got two URLs wrong.

### What happens without it

- **A typo in one URL** and that row fails, or worse, copies into the wrong person's folder.
- **The ShareGate script stops** on a row whose destination folder doesn't exist yet. It needs the folder first.
- **Archived OneDrives fail** with access errors. A OneDrive in Microsoft 365 Archive can't be read by any tool until it's reactivated, and reactivation can take hours and may cost money.
- **Leavers get missed.** People who left a year ago, whose accounts were deleted but whose OneDrives are still held, aren't on anyone's list.

### How it decides

| Who | How it's found | Where it goes |
| --- | --- | --- |
| Named people (`-Users kim, pat@fabrikam.com`) | Their OneDrive in the source, by owner or URL | Their new OneDrive if they're in `-MappingFile`, otherwise a folder in the archive library |
| Leavers (`-Leavers`) | OneDrive owners who are unlicensed, blocked from signing in, or no longer in the directory (Graph) | A folder in the archive library |
| Archived or locked OneDrives | `ArchiveStatus` isn't `NotArchived`, or `LockState` is `NoAccess` | Row marked **Skip** with the reason |
| Two people with the same folder name | Same user name on two domains | Both marked **Skip**. Edit the folder name and set Status back to Copy |

### How to run it

```powershell
# Leavers into one archive library, folders created (preview first)
.\New-DriveCopyList.ps1 -SourceAdminUrl https://fabrikam-admin.sharepoint.com -Leavers `
    -DestinationSite https://contoso.sharepoint.com/sites/Archive -DestinationLibrary 'Former Staff OneDrive' -CreateFolders -WhatIf
.\New-DriveCopyList.ps1 -SourceAdminUrl https://fabrikam-admin.sharepoint.com -Leavers `
    -DestinationSite https://contoso.sharepoint.com/sites/Archive -DestinationLibrary 'Former Staff OneDrive' -CreateFolders

# Current staff into their new OneDrives (rebuild 02b's mapping after cutover first)
.\New-DriveCopyList.ps1 -SourceAdminUrl https://fabrikam-admin.sharepoint.com -Users alex, sam `
    -MappingFile .\UserMapping.csv -DestinationAdminUrl https://contoso-admin.sharepoint.com

# Then copy. ShareGate (Windows PowerShell 5.1):
.\B-ShareGate\Copy-DrivesWithShareGate.ps1 -CsvPath .\DriveCopyList-<date>.csv -SourceAdminUrl https://fabrikam-admin.sharepoint.com `
    -DestinationAdminUrl https://contoso-admin.sharepoint.com -MappingFile .\UserMapping.sgum
# Or no budget (PowerShell 7), two passes:
.\C-No-budget\Copy-DriveViaGraph.ps1 -Download -TenantId <source-guid> -LocalPath C:\T2T -CsvPath .\DriveCopyList-<date>.csv
.\C-No-budget\Copy-DriveViaGraph.ps1 -Upload -TenantId <dest-guid> -LocalPath C:\T2T -CsvPath .\DriveCopyList-<date>.csv
```

Open the CSV before you copy. Every Skip has a Reason. Fix it or accept it, then run the copy.

### Three ways to do it

| Route | How |
| --- | --- |
| Paid tool | ShareGate: copy one OneDrive at a time in the app, or the whole list with its PowerShell module and this CSV (option B) |
| Microsoft's own | Native cross-tenant OneDrive migration moves a OneDrive to the person's new OneDrive. It has no "into a folder of a library" mode, so leavers still need a tool, option C, or a temporary licence |
| By hand | Make yourself site collection admin of each OneDrive, sync or download it, and upload it into the person's folder. Fine for three leavers, not thirty |

### Lessons

- **Trap:** the sign-in name changes at cutover, and so does every destination OneDrive URL. Rebuild the mapping with [02b](../02b-UserMapping/) after cutover, then build this list, then rebuild the ShareGate `.sgum`. Do it in that order before any OneDrive copy or incremental.
- **Tip:** use the user name for folder names, not the display name. It's unique, it survives name changes, and it matches what people search for.
- **Trap:** unlicensed OneDrives don't wait forever. Depending on the tenant's settings they go read-only, into Microsoft 365 Archive, or get deleted. Run `-Leavers` early in the project so you know what you're dealing with.
- **Tip:** you need site collection admin on every source OneDrive for any copy route. `Set-SPOUser -Site <url> -LoginName <you> -IsSiteCollectionAdmin $true` for each row, and remove it afterwards.

## Disclosure

I'm not sponsored by ShareGate or anyone else. I used ShareGate on the migration this toolkit came from because it fit the job, and I've written down exactly what it did well and where it tripped me up.
