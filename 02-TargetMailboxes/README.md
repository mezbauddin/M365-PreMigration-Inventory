# New-TargetMailboxes

**TL;DR:** A PowerShell script that builds the destination mailboxes for a Microsoft 365 tenant-to-tenant migration from a source inventory. Stage every mailbox as an unlicensed shared mailbox so you can start copying mail on day one. Then license, convert and archive them as the licences and the cutover arrive. It's safe to re-run: each run only does what's still missing.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 2.0.1 |
| Released | 7 November 2026 |
| Licence | MIT |

Part 2 of my [M365 tenant-to-tenant migration toolkit](../README.md). Part 1 is [Invoke-SourceTenantInventory](../01-SourceInventory/), which produces the `04-Mailboxes.csv` this script reads.

## Why

Migration tools copy content. They don't decide what the destination should look like. Before any migration tool can copy a mailbox, something has to exist on the other side, with the right type, the right licence and an archive if the source has one.

Waiting for licences, or for the cutover date, holds up the copy. A shared mailbox needs no licence. So build everything as shared first, start the copy, and convert the people later. Mail lands early, and nobody can sign in to the new tenant before you're ready.

## How it works

| Step | Command | What it does |
| --- | --- | --- |
| 1 | `-Plan` | Reads the inventory and writes `TargetBuildPlan.csv`. Changes nothing |
| 2 | `-Apply -StageOnly` | Creates every mailbox as an unlicensed shared mailbox. Start your mail copy |
| 3 | `-Apply -SharedMailboxesOnly` | Licences arrive: licenses and archives the shared mailboxes only |
| 4 | `-Apply` | Cutover: converts people to user mailboxes, sets temporary passwords, licenses, enables archives |

Run step 4 twice, about 15 minutes apart. Archives have to exist before auto-expanding archives can be switched on.

## Requirements

- Global Administrator in the destination tenant (for licences, passwords and consent)
- PowerShell 7 with:
  - `ExchangeOnlineManagement` 3.x
  - `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users` and `Microsoft.Graph.Identity.DirectoryManagement`
- `04-Mailboxes.csv` from [Invoke-SourceTenantInventory](../01-SourceInventory/)

```powershell
Install-Module ExchangeOnlineManagement, Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser
```

## Usage

```powershell
# 1. Plan. Leave out admin accounts and anyone who already exists in the destination.
.\New-TargetMailboxes.ps1 -Plan -TargetDomain contoso.onmicrosoft.com -Exclude admin@fabrikam.com -TreatAsShared leaver@fabrikam.com

# 2. Stage everything as shared mailboxes
.\New-TargetMailboxes.ps1 -Apply -TargetDomain contoso.onmicrosoft.com -UsageLocation GB -StageOnly -WhatIf
.\New-TargetMailboxes.ps1 -Apply -TargetDomain contoso.onmicrosoft.com -UsageLocation GB -StageOnly

# 3. Licences arrived: shared mailboxes only
.\New-TargetMailboxes.ps1 -Apply -TargetDomain contoso.onmicrosoft.com -UsageLocation GB -SharedMailboxesOnly

# 4. Cutover: everyone
.\New-TargetMailboxes.ps1 -Apply -TargetDomain contoso.onmicrosoft.com -UsageLocation GB
```

Add `-StagingSmtpDomain contoso.com` to any `-Apply` to give staged mailboxes a primary address on a domain the destination already owns. The onmicrosoft.com address stays as a proxy, so your migration tool's mapping keeps working.

## What the plan decides

| Source | Planned as | Licence |
| --- | --- | --- |
| User mailbox | User | `-UserLicenceSku` (default `SPB`, Business Premium) |
| User mailbox listed in `-TreatAsShared` | Shared | As below |
| Shared mailbox with an archive, or near `-SharedQuotaGB` | Shared | `-SharedLicenceSku` (default `EXCHANGEENTERPRISE`, Exchange Online Plan 2) |
| Other shared mailboxes | Shared | None |
| Room, equipment or scheduling mailbox | Excluded | Build these by hand. They can't be converted |

Mailboxes on the main source domain keep their alias. Mailboxes on other domains become `alias-domainname`, so `info@fabrikam.com` and `info@legacy.co.uk` don't collide. Archives at or above `-AutoExpandThresholdGB` (default 90) get auto-expanding archiving.

Open the CSV before you apply it. Two things to look for:

- **People who already exist in the destination.** Set `Include` to `False`, and map them by hand in your migration tool. Otherwise you get a duplicate mailbox that takes a licence.
- **Aliases already used in the destination.** `docs` or `marketing` often are. Change `Alias` and `TargetAddress`, e.g. to `fab-docs`.

## Output

| File | Contents |
| --- | --- |
| TargetBuildPlan.csv | The plan. Edit `Include`, `Alias`, `TargetAddress`, `Kind` and `Licence` freely |
| TargetBuild-Status-*.csv | What each run did, per mailbox, with Graph's real error text in `Note` |
| TargetBuild-Passwords-*.csv | Temporary passwords. Sensitive. Move it somewhere secure and delete the local copy |

## Things I hit so you don't have to

- **`Set-MgUserLicense` and the empty `-RemoveLicenses`.** The Graph SDK drops the empty array and Graph returns 400, "missing removeLicenses". The script calls the `assignLicense` API directly.
- **Passwords need their own permission.** `User.ReadWrite.All` isn't enough to set a password. The script asks for `User-PasswordProfile.ReadWrite.All` as well. If you've connected before, run `Disconnect-MgGraph` first so you get the consent prompt.
- **Set the sign-in name on shared mailboxes.** `New-Mailbox -Shared` without `-MicrosoftOnlineServicesID` puts the sign-in name on the tenant's **default** domain, not the address you gave it. In a destination that's usually the customer's real domain. It can clash with existing users, it's what people end up signing in with when you convert them, and it sets their OneDrive URL. I found out when OneDrives appeared as `name_theirdomain_com`.
- **Usage location before licence.** No usage location, no licence. The script sets it on its own, so a failed password step can't block licensing.
- **New archives aren't instant.** `ArchiveStatus` says Active straight away, but the archive itself can take a while. `Get-EXOMailboxStatistics -Archive` returning nothing means it isn't ready yet.
- **Archive copies often need EWS, and the EWS allow list changed in October 2026.** Main mail can copy fine while every archive fails with 'Unauthorized'. [Test-EwsArchiveReadiness](../03-EwsArchiveReadiness/) checks and fixes the causes in both tenants.
- **Some tools copy only email into a shared mailbox.** The one I used skipped calendars and contacts until the mailbox was a user mailbox. Check yours, and run an incremental after step 4 if needed.

## Notes

- `-Apply` checks that the signed-in tenant owns `-TargetDomain` before it changes anything. Run it against the wrong tenant and it stops.
- Everything supports `-WhatIf`. Use it every time you change the plan.
- Licences are only assigned from free seats. Short on seats, and the row shows `PENDING: no free <SKU>` until you buy more and re-run.

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 2.0.1 | 2026-11-08 | Shared mailboxes get their sign-in name set to `TargetAddress`. Without it Exchange uses the tenant's default domain. Warns when a sign-in name differs, and the passwords file shows the real one |
| 2.0.0 | 2026-11-07 | First public release. Generic parameters replace tenant-specific settings. Adds `-StagingSmtpDomain`, `-ResetPasswords`, `-TreatAsShared` and `-Exclude` |
| 1.2.2 | 2026-11-07 | Password permission added. Usage location and licence no longer depend on the password step. Resumes people converted by an interrupted run. Warnings show Graph's real error text |
| 1.2.1 | 2026-11-07 | Licences assigned through the Graph API, working round the `Set-MgUserLicense` empty-array bug |
| 1.2.0 | 2026-11-07 | `-SharedMailboxesOnly` |
| 1.1.0 | 2026-11-06 | Stage everything as shared mailboxes, convert people later |
| 1.0.0 | 2026-11-06 | First version, used on a live migration |
