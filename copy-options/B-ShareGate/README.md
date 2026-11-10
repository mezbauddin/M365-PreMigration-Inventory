# Option B: ShareGate

**TL;DR:** This is how I copied a small business's mail, archives, Teams, SharePoint and OneDrive between tenants with ShareGate Migrate, in the order I'd do it again. Most of it really is point and click. The parts that weren't are below, with the fix for each, so you don't lose an evening to them like I did.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 2.0.0 |
| Tested | Production, on a small business tenant: mailboxes with large archives, SharePoint, Teams and OneDrive |
| Tool version | ShareGate Migrate 26.11 |
| Script | [`Copy-DrivesWithShareGate.ps1`](Copy-DrivesWithShareGate.ps1) 1.1.0, [`Copy-PeopleOneDrives.ps1`](Copy-PeopleOneDrives.ps1) 1.0.0 |
| Licence | MIT (this page and script; ShareGate itself is a paid product) |

Part of the [copy options](../README.md) in my [M365 tenant-to-tenant migration toolkit](../../README.md). Not sponsored. I chose it because it fit the job.

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 2.0.0 | 2026-11-10 | Adds what happens without it, three ways and a tips and traps summary. Live-job specifics replaced with generic wording; the lessons are unchanged |
| 1.2.0 | 2026-11-10 | `Copy-PeopleOneDrives.ps1` 1.0.0: people's OneDrives end to end (finds real URLs by owner in both tenants, creates missing destination OneDrives, adds you as admin, runs the copy) |
| 1.1.0 | 2026-11-10 | `Copy-DrivesWithShareGate.ps1` 1.1.0 leaves out rows marked Skip by `New-DriveCopyList.ps1` |
| 1.0.0 | 2026-11-09 | First version, written up from a live migration |

## Why ShareGate for this one

- **SharePoint and Teams.** Microsoft's native route doesn't move Teams content, and its SharePoint add-on was Enterprise Agreement only. This customer wasn't on one.
- **Size.** For a small business, setting up native migration (app registration, trust between tenants, identity mapping) would have taken longer than the copy itself.
- **Incrementals.** Copy early, top up before cutover, top up again after. Same job, same mapping.

## What happens without it

Without a tool like this (or option A), a business with real SharePoint and Teams use is left with [option C](../C-No-budget/): PSTs for mail and download-and-upload for files. That means:

- **Versions, authors and permissions are lost.** Every file shows as created by the migration account on the day of the copy.
- **No incrementals.** Mail and files have to freeze while the copy runs, or whatever changed afterwards stays behind.
- **Teams arrive as folders.** Channel files survive; the team structure, membership and tabs have to be rebuilt by hand.

And without the steps on this page, ShareGate itself trips over the gotchas below: archives fail on EWS, shared mailboxes arrive without calendars, members don't come across because the mapping wasn't imported.

## Before you open ShareGate

Run these from the toolkit first. They make the ShareGate steps short:

1. [01 Inventory](../../01-SourceInventory/), so you know sizes, archives and Teams.
2. [02 Target mailboxes](../../02-TargetMailboxes/). ShareGate copies into mailboxes; it doesn't create them.
3. [03 EWS archive readiness](../../03-EwsArchiveReadiness/) in **both** tenants with ShareGate's app ID, if you're copying archives.
4. A user mapping CSV, source address to destination sign-in. The same file drives mail config ([05](../../05-MailConfig/)) and cutover ([07](../../07-DomainCutover/)).

## 1. Mailboxes

**Copy mailboxes**, source tenant to destination tenant.

- **Date range:** Everything. Tick deleted items and junk if the business wants them.
- **Archives:** tick **Copy In-Place Archive**. If you're not ready, leave it off and tick it on a later incremental pass. Already-copied mail is skipped.
- **Mapping:** check every row. Display-name matching can pick the wrong mailbox when two people share a name, or when a mailbox already exists in the destination. Export the mapping and keep it.
- **Incrementals:** same PC, same ShareGate install, same mapping. That's how it knows what it has already copied. Don't clear its history.

**Gotchas I hit:**

| What happened | Why | Fix |
| --- | --- | --- |
| Main mail copied, every archive failed, 'Unable to access source archive mailbox: Unauthorized' | Archives go over EWS. The source still had an old EWS app-name rule, and the new app ID allow list was missing | [03 EWS archive readiness](../../03-EwsArchiveReadiness/) in both tenants, then wait and retry |
| Shared mailbox destinations got email but no calendar or contacts | The tool copies only email into a shared mailbox | Convert people to user mailboxes ([02](../../02-TargetMailboxes/)), then run an incremental |
| A scheduling mailbox wouldn't convert | Bookings-style mailboxes can't become shared or user mailboxes | Leave it out, or rebuild it by hand |

## 2. Teams

**Copy Teams**, then choose per team:

- **Copy as a new team.** ShareGate creates it. Simple, but if a team with that name already exists you get a second one, with a site URL ending in 2.
- **Copy channels into an existing team.** Use this if you've already created the destination teams, for example so their email addresses land on the right domain.

**Gotchas I hit:**

| What happened | Fix |
| --- | --- |
| The team picker search found nothing for a word in the middle of a team name ('board' in 'Client Onboarding') | Search from the start of the name |
| I forgot to import the user mapping, so members didn't come across | Export members from the source groups and add them with `Add-UnifiedGroupLinks`, mapped through the same CSV. Import the mapping before you start next time |
| One channel showed an error, but its files were all green | The error was on channel settings, not content. Check the details before re-running |

## 3. SharePoint

**Copy content**, one library at a time, into the destination site.

- **Don't drop a whole site onto an existing site.** ShareGate treats that as 'make a subsite'. Expand the source site and select its libraries, then pick the destination site.
- **Subsites** show under a **Sites** node in the source. They're separate jobs.
- **Skip the system libraries:** Form Templates, Master Page Gallery, Site Assets, Site Pages, Style Library, Web Part Gallery.

**Gotchas I hit:**

| What happened | Fix |
| --- | --- |
| 'Copying to a Microsoft 365 site collection requires Custom Scripts to be enabled', then 'not authorized' | Allow custom scripts on the destination site yourself: `Set-SPOSite -Identity <url> -DenyAddAndCustomizePages 0`. It switches itself off again after about a day, so re-run before incrementals |
| 'Strongly recommended to have Site Collection Administrator permissions on your source site' | `Set-SPOUser -Site <url> -LoginName <you> -IsSiteCollectionAdmin $true` on both sides |
| The mapping import wouldn't see my CSV | It only accepts its own `.sgum` format. Convert the CSV (example below) |

Converting a `Source,Destination` CSV into a `.sgum` mapping file:

```powershell
$rows = Import-Csv .\UserMapping.csv
$x = [System.Text.StringBuilder]::new()
[void]$x.AppendLine('<?xml version="1.0" encoding="utf-8"?>')
[void]$x.AppendLine('<UserAndGroupMappings xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">')
[void]$x.AppendLine('  <Mappings>')
foreach ($r in $rows) {
    $s = [Security.SecurityElement]::Escape($r.Source.ToLower())
    $d = [Security.SecurityElement]::Escape($r.Destination.ToLower())
    [void]$x.AppendLine("    <Mapping><Source AccountName=""i:0#.f|membership|$s"" DisplayName=""$s"" PrincipalType=""User"" /><Destination AccountName=""i:0#.f|membership|$d"" DisplayName=""$d"" PrincipalType=""User"" /></Mapping>")
}
[void]$x.AppendLine('  </Mappings>')
[void]$x.AppendLine('</UserAndGroupMappings>')
[IO.File]::WriteAllText("$PWD\UserMapping.sgum", $x.ToString(), [Text.UTF8Encoding]::new($false))
```

## 4. OneDrive

- Create the destination OneDrives first ([04](../../04-TargetOneDrives/)). ShareGate can't copy into one that doesn't exist.
- Copy **after** sign-in names are final, because the OneDrive URL follows the sign-in name ([07](../../07-DomainCutover/)).
- Leavers have no destination OneDrive. Copy theirs into a library on a SharePoint site, one folder per person.

### People's OneDrives after cutover: `Copy-PeopleOneDrives.ps1`

After a cutover, both tenants rename OneDrive URLs in the background, so a guessed URL fails with **CO-001-032** ("The connection to the SharePoint site failed"). Destination OneDrives may not exist yet, and you need site admin on every one. [`Copy-PeopleOneDrives.ps1`](Copy-PeopleOneDrives.ps1) takes user names and your mapping, finds each OneDrive **by owner** in both tenants, skips people who never used OneDrive (nothing to copy), creates missing destination OneDrives, adds you as admin, then runs `Copy-DrivesWithShareGate.ps1`.

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Copy-PeopleOneDrives.ps1 -MappingFile .\Mapping-PostCutover\UserMapping.csv `
    -Users kim, lee, sam -SourceAdminUrl https://fabrikam-admin.sharepoint.com -SourceAdmin admin@fabrikam.onmicrosoft.com `
    -DestinationAdminUrl https://contoso-admin.sharepoint.com -DestinationAdmin admin@contoso.com -SgumFile .\Mapping-PostCutover\UserMapping.sgum
```

- **Trap:** a person with no source OneDrive never used it. There is nothing to copy; their new OneDrive is created when they first sign in.
- **Tip:** with Files On-Demand, most of a person's files exist only in the cloud. Copy their OneDrive **before** you move their laptop, or they sign in to empty folders.

### Many OneDrives at once: `Copy-DrivesWithShareGate.ps1`

The app copies OneDrives one at a time. ShareGate's PowerShell module (Pro and Enterprise plans) doesn't have that limit. [`Copy-DrivesWithShareGate.ps1`](Copy-DrivesWithShareGate.ps1) reads a CSV of source and destination pairs, signs in to each tenant once, and runs `Copy-Content` for every row with your `.sgum` mapping. Each row can land in the person's new OneDrive or in a folder of a library, for example leavers' OneDrives into an archive library.

```text
Source,DestinationSite,DestinationLibrary,DestinationFolder
https://fabrikam-my.sharepoint.com/personal/alex_fabrikam_com,https://contoso-my.sharepoint.com/personal/alex_contoso_com,,
https://fabrikam-my.sharepoint.com/personal/kim_fabrikam_com,https://contoso.sharepoint.com/sites/Archive,Former Staff OneDrive,kim
```

```powershell
# Windows PowerShell 5.1 only. From PowerShell 7, prefix with: powershell.exe -ExecutionPolicy Bypass -File
.\Copy-DrivesWithShareGate.ps1 -CsvPath .\drives.csv -SourceAdminUrl https://fabrikam-admin.sharepoint.com `
    -DestinationAdminUrl https://contoso-admin.sharepoint.com -MappingFile .\UserMapping.sgum -WhatIf   # pre-check
.\Copy-DrivesWithShareGate.ps1 -CsvPath .\drives.csv -SourceAdminUrl https://fabrikam-admin.sharepoint.com `
    -DestinationAdminUrl https://contoso-admin.sharepoint.com -MappingFile .\UserMapping.sgum
# Later, top up
.\Copy-DrivesWithShareGate.ps1 ... -Incremental
```

Destination folders must exist first. You need site collection admin on every source OneDrive, and the destination OneDrives must exist ([04](../../04-TargetOneDrives/)).

**Don't write the CSV by hand.** [`New-DriveCopyList.ps1`](../New-DriveCopyList.ps1) in the copy-options folder builds it: leavers (unlicensed, blocked or deleted owners) into one folder each, current staff into their new OneDrive through your mapping, archived OneDrives marked Skip. With `-CreateFolders` it makes the folders too. Version 1.1.0 of `Copy-DrivesWithShareGate.ps1` leaves out rows marked Skip.

```powershell
# PowerShell 7: build the list and the folders
..\New-DriveCopyList.ps1 -SourceAdminUrl https://fabrikam-admin.sharepoint.com -Leavers `
    -DestinationSite https://contoso.sharepoint.com/sites/Archive -DestinationLibrary 'Former Staff OneDrive' -CreateFolders
# Windows PowerShell 5.1: copy
.\Copy-DrivesWithShareGate.ps1 -CsvPath .\DriveCopyList-<date>.csv -SourceAdminUrl https://fabrikam-admin.sharepoint.com `
    -DestinationAdminUrl https://contoso-admin.sharepoint.com -MappingFile .\UserMapping.sgum
```

**Trap:** sign-in names change at cutover. Rebuild the mapping ([02b](../../02b-UserMapping/)) and the `.sgum` afterwards, then rebuild this list, before any OneDrive copy or incremental. An incremental with the old `.sgum` maps nobody.

**Trap:** ShareGate's module only loads in Windows PowerShell 5.1. If you've just run Graph or Exchange scripts in 7, open a separate 5.1 window rather than calling it from the same session. [`tools/Test-ModuleHealth.ps1`](../../tools/) tells you which window can load what.

## 5. Cutover and after

1. Final incremental on mail, Teams and sites, just before [domain cutover](../../07-DomainCutover/).
2. After cutover: OneDrive, then one more incremental on everything.
3. Re-run the custom scripts command on destination sites before any SharePoint incremental.

## What I'd do differently

- Import the user mapping into **every** job before clicking Start.
- Run the EWS readiness check in both tenants on day one, not after the first archive failure.
- Create Teams and group-connected sites myself, so their email addresses and URLs are what I want, then copy channels into them.

## Three ways to do it

| Route | How |
| --- | --- |
| Paid tool | This page: ShareGate Migrate for mail, archives, Teams, SharePoint and OneDrive, with [`Copy-DrivesWithShareGate.ps1`](Copy-DrivesWithShareGate.ps1) for many OneDrives at once. Other migration suites work the same way: build the destination, map users, copy, top up |
| Microsoft's own | [A. Microsoft native](../A-Microsoft-native/) for mailboxes and OneDrive. Nothing native for shared Teams content, and SharePoint only with the Enterprise Agreement add-on. A common split is native for mail and ShareGate for SharePoint and Teams |
| By hand | [C. No budget](../C-No-budget/): PST export and import for mail, Graph or the sync client for files. Fine for a handful of people with little shared data |

## Tips and traps

- **Trap:** archives copy over EWS. Run [03 EWS archive readiness](../../03-EwsArchiveReadiness/) in both tenants before the first mail job, and finish big archives before EWS ends on 1 April 2027.
- **Trap:** a shared mailbox destination gets email only. Convert people to user mailboxes ([02](../../02-TargetMailboxes/)) and run an incremental.
- **Trap:** the user mapping isn't remembered between job types. Import the `.sgum` into every mail, Teams, SharePoint and OneDrive job, and rebuild it after cutover.
- **Trap:** custom scripts on a destination site switch themselves off after about a day. Re-run `Set-SPOSite -DenyAddAndCustomizePages 0` before every SharePoint incremental.
- **Tip:** use the same PC and the same ShareGate install for every incremental. Its local history is how it knows what's already copied.
- **Tip:** export each job's report and keep it with the toolkit's CSVs. ShareGate's report is your evidence of what moved; [tools/Get-MigrationReport.ps1](../../tools/) covers the toolkit's side.
