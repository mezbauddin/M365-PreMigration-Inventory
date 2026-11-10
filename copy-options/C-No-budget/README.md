# Option C: no budget

**TL;DR:** You can move a small tenant with nothing but what Microsoft 365 already includes: export mail to PST files (eDiscovery or Outlook), upload them with AzCopy and import them with Microsoft's free PST import service, and move files with Microsoft Graph or the sync client. [`New-PstImportMapping.ps1`](New-PstImportMapping.ps1) writes the import mapping CSV so no PST lands in the wrong mailbox, and [`Copy-DriveViaGraph.ps1`](Copy-DriveViaGraph.ps1) copies OneDrives and libraries. It works for a handful of people with little shared data. It's slow, you lose versions and most metadata, and there's no incremental pass, so plan a short freeze. Past about ten people, or with real SharePoint use, the time you spend will cost more than a tool.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 2.0.0 |
| Status | Written from Microsoft's documentation and standard practice. Both scripts tested offline; the PST path hasn't yet been run end to end for this toolkit |
| Licence | MIT |

Part of the [copy options](../README.md) in my [M365 tenant-to-tenant migration toolkit](../../README.md).

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 2.0.0 | 2026-11-10 | Full no-budget mail path (export, upload, mapping, import job) and its limits. Ships `New-PstImportMapping.ps1` 1.0.0. Adds what happens without it, three ways, tips and traps |
| 1.1.0 | 2026-11-10 | `Copy-DriveViaGraph.ps1` 1.1.0 with `-CsvPath` for the list from `New-DriveCopyList.ps1` |
| 1.0.0 | 2026-11-09 | First version |

## Scripts

| Script | Version | Purpose |
| --- | --- | --- |
| [`New-PstImportMapping.ps1`](New-PstImportMapping.ps1) | 1.0.0 | Scans a folder of PSTs, matches each to the user mapping, writes the PST import mapping CSV and a review list of anything it couldn't match. Prints the AzCopy command. Reads local files only |
| [`Copy-DriveViaGraph.ps1`](Copy-DriveViaGraph.ps1) | 1.1.0 | OneDrive or library copy between tenants with only Microsoft Graph. `-Download` from the source to a local folder, then `-Upload` into the destination, keeping folders and file dates. Skips what's already done, so re-runs are safe. Large files go up in 10 MB chunks. `-CsvPath` takes the list from [`New-DriveCopyList.ps1`](../New-DriveCopyList.ps1) |

## Why

Some migrations have no budget line at all: a five-person office, a charity, a company being wound into its new parent. The tools in options A and B cost money per user or per month. Everything on this page uses services that come with the licences the business already has.

## What happens without it

If there's no budget and no plan, the usual fallback is "everyone exports their own PST and imports it in Outlook". What that gets you:

- **Half the mail.** People export the mailbox but not the archive, or the archive but not the shared mailbox they rely on.
- **Imports into the wrong place.** A PST opened in Outlook and dragged across lands in whichever mailbox was selected. Do that for a shared mailbox and the mail is now in someone's personal one.
- **No record.** Nobody can say afterwards which mailboxes were moved and which weren't.

The PST import service with a mapping CSV fixes all three: every PST goes to a named mailbox or archive, and the CSV is your record.

## What you keep and what you lose

| | Kept | Lost or manual |
| --- | --- | --- |
| Mail | Messages, folders, calendar and contacts, in the PST | Rules, some categories, delegate settings, Teams chat. Archives need their own PST |
| Files | Files, folders and file dates | Version history, created and modified by (they become you), sharing links, unique permissions |
| Teams | Channel files (they live in SharePoint) | Channel conversations, tabs, apps, Planner |

## Mail: the full no-budget path

### 1. Export PSTs from the source

Two ways. Use eDiscovery for more than a couple of people.

| Route | How | Watch for |
| --- | --- | --- |
| eDiscovery (Purview) | Purview > eDiscovery > create a case > a search with the mailboxes as data sources, no conditions > Export, as PST, one PST per mailbox. Download the export | You need the eDiscovery Manager role. Check whether the archive came out inside the mailbox PST or as its own file |
| Outlook (classic) | Signed in as the person, or with Full Access: File > Open & Export > Import/Export > Export to a file > Outlook Data File (.pst), tick Include subfolders | One mailbox at a time. The archive only comes out if it's in the profile. The new Outlook can't export PSTs |

### 2. Name the files

Rename each PST to one of:

| File name | Goes into |
| --- | --- |
| `alex.pst` or `alex@fabrikam.com.pst` | Alex's mailbox |
| `alex_archive.pst` or `alex@fabrikam.com_archive.pst` | Alex's archive |

`alex` is the part of the **source** address before the @. Use the full address when two source domains share an alias (`info@fabrikam.com` and `info@legacy.co.uk`); the script flags a short name that matches both. No spaces or special characters.

### 3. Build the mapping CSV

```powershell
.\New-PstImportMapping.ps1 -PstFolder C:\PST -MappingFile .\UserMapping.csv
# PSTs into a folder called Imported instead of merged into the mailbox's own folders
.\New-PstImportMapping.ps1 -PstFolder C:\PST -MappingFile .\UserMapping.csv -TargetRootFolder /Imported
```

It writes `PstImportMapping-<date>.csv` in the exact format the import service wants:

```text
Workload,FilePath,Name,Mailbox,IsArchive,TargetRootFolder,ContentCodePage,SPFileContainer,SPManifestContainer,SPSiteUrl
Exchange,,alex.pst,alex@contoso.com,FALSE,/,,,,
Exchange,,alex_archive.pst,alex@contoso.com,TRUE,/,,,,
```

and `PstImport-Review-<date>.csv`, one row per PST: Matched, Unmatched or Ambiguous, with the reason, the size, and a note when two PSTs go to the same place or a file is over 20 GB. Unmatched and Ambiguous PSTs are left out of the mapping. Fix them (rename, or add a mapping row) and re-run.

| `-TargetRootFolder` | Result |
| --- | --- |
| `/` (default) | Folders merged into the mailbox's own: the PST's Inbox into the Inbox |
| `/Imported` | Everything under a folder called Imported |
| `''` (blank column) | The service makes a folder called Imported |

### 4. Upload with AzCopy

In the **destination**: Purview > Data lifecycle management > Microsoft 365 > Import > New import job > Upload your data > **Show network upload SAS URL**. Copy it. Then, with AzCopy v10:

```powershell
azcopy copy "C:\PST\*" "<SAS URL>" --recursive
```

The script prints this with your folder filled in. With `-FilePath Batch1`, put `/Batch1` into the SAS URL just before the `?` and the mapping's FilePath column says `Batch1`.

**Tip:** the SAS URL comes from Purview > Data lifecycle management > Import, and it's a key to the upload location. Keep it out of tickets, chat and this repo.

**Trap:** `"C:\PST"` without the `\*` uploads the folder itself, so every PST sits under `PST/` and every row's FilePath is wrong. The import then says it can't find the files.

### 5. Run the import job

Same wizard: tick that you've finished uploading and have the mapping file, upload `PstImportMapping-<date>.csv`, **Validate**, fix anything it rejects, save. Once the service has analysed the PSTs, choose to import everything (or filter by age), and start it.

### 6. Freeze, then a final pass

There's no incremental. Either accept a cut-off (mail after the export date stays behind, or arrives in the new tenant after cutover anyway), or export once more for a short window just before cutover and import that with `-TargetRootFolder /` too.

### Limits and traps for the mail path

- **Trap: the account needs the Mailbox Import Export role** in the destination's Exchange admin centre, and it isn't in any role group by default. Assign it to yourself a day early; it can take a while to apply.
- **Trap: the destination mailbox or archive must exist first.** Build mailboxes with [02](../../02-TargetMailboxes/), and enable archives before importing `_archive` PSTs. A row for a missing archive fails.
- **Trap: retention hold.** The service turns on retention hold for each mailbox it imports into, so retention policies don't delete old imported mail straight away. Once people have checked their mail, decide whether to turn it off (`Set-Mailbox <address> -RetentionHoldEnabled $false`). Otherwise policies you set later quietly do nothing.
- **Trap: import twice and you get everything twice.** There's no de-duplication. The review CSV flags two PSTs going to the same mailbox; check it before every job.
- **Tip:** keep each PST under 20 GB. Bigger ones import, but slowly, and a failure means starting that file again. Single items over 150 MB are skipped.
- **Tip:** uploads stay in Microsoft's storage for a limited time, so run the import job within days of uploading, not weeks.
- **Tip:** imported mail still carries the old internal addresses. The X500 step in [07](../../07-DomainCutover/) matters more here, or replies to old messages bounce.

## Files: Graph or the sync client

```powershell
# Source: download each OneDrive into C:\T2T\<name>
.\Copy-DriveViaGraph.ps1 -Download -TenantId <source-guid> -LocalPath C:\T2T -SourceSiteUrl `
    https://fabrikam-my.sharepoint.com/personal/alex_fabrikam_com, https://fabrikam-my.sharepoint.com/personal/kim_fabrikam_com

# Destination: upload each C:\T2T\<name> into a folder of the same name in a library
.\Copy-DriveViaGraph.ps1 -Upload -TenantId <dest-guid> -LocalPath C:\T2T `
    -DestinationSiteUrl https://contoso.sharepoint.com/sites/Archive -DestinationLibrary 'Former Staff OneDrive'
```

```powershell
# Or from the copy list (see ../New-DriveCopyList.ps1): leavers into one folder each, staff into their new OneDrive
.\Copy-DriveViaGraph.ps1 -Download -TenantId <source-guid> -LocalPath C:\T2T -CsvPath ..\DriveCopyList-<date>.csv
.\Copy-DriveViaGraph.ps1 -Upload -TenantId <dest-guid> -LocalPath C:\T2T -CsvPath ..\DriveCopyList-<date>.csv
```

**Tip:** with `-CsvPath`, a row with a DestinationFolder downloads into `C:\T2T\<folder>` and uploads into that folder of the library. A row without one (OneDrive to OneDrive) uses the URL's last segment locally and uploads into the root of the person's new OneDrive.

It's a good fit for leavers' OneDrives: small, nobody working in them, no need for version history. Archived OneDrives (unlicensed for months) can't be read until they're reactivated.

- **Sync client instead:** sync the source OneDrive or library, copy the folder, sync it into the destination. Or have each person do their own. Recreate permissions by hand afterwards.
- **Keep paths short.** Long paths that worked in the source can fail on the way back in.

## Three ways to do it

| Route | How |
| --- | --- |
| Paid tool | [B. ShareGate](../B-ShareGate/) or similar: mail, archives, files with versions and authors, incrementals. If the budget appears, this is the first thing to buy |
| Microsoft's own | This page: eDiscovery export, PST import service, Graph. Or [A. native migration](../A-Microsoft-native/) if the per-user add-on fits the budget better than your time does |
| By hand | Each person exports a PST in Outlook and imports it in the new tenant's Outlook, and copies their own files. Fine for two or three people with tidy mailboxes. Keep a checklist of who has done what |

## Everything else from the toolkit still applies

Build the destination ([02](../../02-TargetMailboxes/)), rebuild permissions, lists and forwarding ([05](../../05-MailConfig/)), and cut over the domain ([07](../../07-DomainCutover/)) exactly as you would with a tool.
