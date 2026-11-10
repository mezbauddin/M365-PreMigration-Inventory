# Tools

**TL;DR:** Two helpers that sit beside the numbered parts. `Test-ModuleHealth.ps1` checks your PowerShell modules for the version clashes that stop a script half way through a cutover, and can fix the Microsoft Graph ones. `Get-MigrationReport.ps1` reads every report CSV the toolkit wrote and shows you, in one table, what still needs a look. Neither signs in to a tenant.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Licence | MIT |

| Script | Version | Changes anything? |
| --- | --- | --- |
| [`Test-ModuleHealth.ps1`](Test-ModuleHealth.ps1) | 1.0.0 | No. Only with `-Fix` (installs Graph modules for the current user) |
| [`Get-MigrationReport.ps1`](Get-MigrationReport.ps1) | 1.0.0 | No. Writes an HTML file with `-Html` |

Part of my [M365 tenant-to-tenant migration toolkit](../README.md).

## Test-ModuleHealth

### Why you need it

On a live cutover the destination scripts stopped with this, part way through:

```text
Could not load file or assembly 'Microsoft.Graph.Authentication, Version=2.26.1.0' ...
Assembly with same name is already loaded
```

The PC had `Microsoft.Graph.Authentication` 2.41.2 and an older `Microsoft.Graph.Identity.DirectoryManagement` 2.26.1. PowerShell loaded the new Authentication module first, then the old submodule asked for its own, older Authentication, and .NET won't load two versions of the same assembly in one window. The fix was to install every `Microsoft.Graph.*` module at one version and open a fresh window. This script finds that before the night, not during it.

### What happens without it

- A script fails after it has already changed half the tenant, and you're reinstalling modules at 10pm.
- You fix it in one window, but the window you're working in already loaded the old assembly, so it still fails. Only a fresh window clears it.
- Exchange Online and Graph share `Microsoft.Identity.Client`. Load one, then the other at an incompatible version, and sign-in breaks in odd ways.

### What it checks

| Check | WARN when |
| --- | --- |
| `Microsoft.Graph.Authentication` | Not installed, or more than one version installed |
| Every other `Microsoft.Graph.*` | Its newest version isn't the newest Authentication version. This is the clash above |
| `ExchangeOnlineManagement` | More than one version installed |
| `Microsoft.Online.SharePoint.PowerShell`, `PnP.PowerShell`, `MicrosoftTeams` | Listed for information |
| PowerShell edition and ShareGate | The `Sharegate` module is present but you're in PowerShell 7. It only loads in Windows PowerShell 5.1 |
| This window | Exchange Online or Graph assemblies are already loaded. Open a fresh window |

### How to run it

```powershell
# Read-only
.\Test-ModuleHealth.ps1

# See what -Fix would install, then do it
.\Test-ModuleHealth.ps1 -Fix -WhatIf
.\Test-ModuleHealth.ps1 -Fix -Confirm:$false

# Keep the table
.\Test-ModuleHealth.ps1 -OutFile .\ModuleHealth.csv
```

After `-Fix`, close **every** PowerShell window and open a new one. `07 Test-CutoverReadiness.ps1` runs this for you as part of the go/no-go.

### Three ways to do it

| Route | How |
| --- | --- |
| Paid tool | ShareGate installs and loads its own module, but only in Windows PowerShell 5.1. Keep it in its own 5.1 window and run the toolkit in 7 |
| Microsoft's own | `Get-InstalledModule Microsoft.Graph* \| Select Name, Version`, then `Update-Module Microsoft.Graph` (the roll-up module) so every submodule moves together |
| By hand | `Uninstall-Module <name> -AllVersions`, then `Install-Module <name> -RequiredVersion <one version> -Scope CurrentUser` for each module the scripts list under Requires |

### Lessons

- **Tip:** install the Graph modules you need by name and at one version, rather than the full `Microsoft.Graph` roll-up. It's smaller and there's less to drift.
- **Trap:** `Update-Module Microsoft.Graph.Authentication` on its own is how the clash starts. Update all of them, or none.
- **Trap:** a window that already ran `Connect-ExchangeOnline` has loaded its identity library. Running a Graph script in the same window afterwards can fail even when versions match. One task, one fresh window.
- **Tip:** ShareGate scripts go in `powershell.exe` (5.1). Everything else in `pwsh` (7). Two windows, side by side, labelled.

## Get-MigrationReport

### Why you need it

Every toolkit script writes a report: `MailConfig-Import-*`, `SwitchTargetAddresses-*`, `RemoveSourceDomain-*`, `DriveCopy-ShareGate-*`, `CutoverReadiness-*` and more. Each has a few rows marked Skipped, Conflict, Problem or Blocker. On the night you need one answer: what's left? This reads them all and tells you.

### What happens without it

- A Conflict in one CSV (an address already used by a guest) goes unread, and a person's mail bounces for a week.
- You open twelve CSVs in Excel, filter each one, and miss the one written to a subfolder.
- The business owner asks "is it done?" and you don't have a one-page answer.

### What it does

- Finds the newest CSV of each known type in `-Folder` (or every one with `-All`), plus any other CSV with a `Status` or `Result` column. Files with "Password" in the name are never read.
- Counts rows by Status or Result.
- Prints every row that isn't Done, Exists, OK, WhatIf, Successful, Pass, Created, Added, Set or similar, with the useful columns picked out (Object, Source, Status, Detail, Reason, Fix and so on). ShareGate rows with errors are shown even when the result says Successful.
- `TargetDomainDns-*.csv` has no status column, so all its rows are shown: they're the records to publish.
- `-Html` writes `MigrationReport-<date>.html`, one self-contained page, no external files.

### How to run it

```powershell
# In the folder you've been running the scripts from
.\Get-MigrationReport.ps1

# Include subfolders (Copy-DriveViaGraph writes into its -LocalPath) and write the HTML
.\Get-MigrationReport.ps1 -Folder C:\T2T -Recurse -Html

# Every report, not just the newest of each kind
.\Get-MigrationReport.ps1 -All -MaxRows 500
```

### Three ways to do it

| Route | How |
| --- | --- |
| Paid tool | ShareGate's Tasks view shows its own jobs, item by item. It knows nothing about your Exchange, DNS or domain steps. Use both |
| Microsoft's own | Microsoft's native migration batches report in the Exchange admin centre and `Get-MigrationUser`. Again, only the copy |
| By hand | Open each CSV in Excel, filter Status to anything but Done, Exists and WhatIf. Fine for three files, painful for twelve |

### Lessons

- **Tip:** run it after every step on cutover night, not just at the end. A Blocker found at step 4 is cheaper than one found at step 10.
- **Trap:** `-WhatIf` runs used to skip writing their reports, so the dry run left no evidence. Every toolkit script now writes its report even under `-WhatIf` (`-WhatIf:$false` on the export), and this tool treats WhatIf rows as fine.
- **Tip:** attach the HTML to the handover. It's the closest thing to a straight answer when someone asks what didn't move.
- **Trap:** the HTML holds names and addresses. Treat it like the CSVs.

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 1.0.0 | 2026-11-10 | First release: Test-ModuleHealth and Get-MigrationReport, from lessons on a live cutover |
