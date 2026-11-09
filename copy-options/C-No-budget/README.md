# Option C: no budget

**TL;DR:** You can move a small tenant with nothing but what Microsoft 365 already includes: export mail to PST files with eDiscovery, import them with Microsoft's free PST import service, and move files by downloading and re-uploading them. It works for a handful of people with little shared data. It's slow, you lose versions and most metadata, and there's no incremental pass, so plan a short freeze. Past about ten people, or with real SharePoint use, the time you spend will cost more than a tool.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Status | Written up from standard practice. Scripts follow once lab-tested |

Part of the [copy options](../README.md) in my [M365 tenant-to-tenant migration toolkit](../../README.md).

## What you keep and what you lose

| | Kept | Lost or manual |
| --- | --- | --- |
| Mail | Messages, folders, calendar and contacts, in the PST | Rules, some categories, delegate settings. Archives must be exported separately |
| Files | Files and folders | Version history, created and modified by, sharing links, unique permissions |
| Teams | Channel files (they live in SharePoint) | Channel conversations, tabs, apps, Planner |

## Mail: eDiscovery export, then PST import

1. **Source: export.** Microsoft Purview, eDiscovery, a search per mailbox (or all of them), then export to PST. Include archives if you need them; they come out as separate PSTs.
2. **Download** with the export tool. Big tenants mean big downloads, so give it time and disk space.
3. **Destination: PST import.** Purview, Data lifecycle management, Import, then **upload** your PSTs to the Azure storage location it gives you (AzCopy). Map each PST to a mailbox, or to its archive, in the mapping CSV.
4. **Freeze, then a final export.** There's no incremental. Either accept a cut-off date, or export once more for a short window just before cutover and import that too.

## Files: download and re-upload

- **OneDrive:** sync the source OneDrive, copy the folder, then sync it into the destination OneDrive. Or have each person do their own.
- **SharePoint:** sync each library and copy it into a library on the destination site. Recreate permissions by hand afterwards.
- **Keep paths short.** Long paths that worked in the source can fail on the way back in.

## Everything else from the toolkit still applies

Build the destination ([02](../../02-TargetMailboxes/)), rebuild permissions, lists and forwarding ([05](../../05-MailConfig/)), and cut over the domain ([07](../../07-DomainCutover/)) exactly as you would with a tool. The X500 step in 07 matters even more here: imported mail still carries the old internal addresses.

## Scripts planned for this folder

| Script | Purpose |
| --- | --- |
| `New-PstImportMapping.ps1` | Builds the PST import mapping CSV from the toolkit's user mapping |
| `Copy-LibraryWithPnP.ps1` | Library to library copy between tenants through a local folder, with a report of what didn't make it |
