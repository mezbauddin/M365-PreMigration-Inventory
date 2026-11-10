# New-TargetMailboxes

**TL;DR:** A PowerShell script that builds the destination mailboxes for a Microsoft 365 tenant-to-tenant migration from a source inventory. Stage every mailbox as an unlicensed shared mailbox so you can start copying mail on day one. Then license, convert and archive them as the licences and the cutover arrive. It's safe to re-run: each run only does what's still missing. Rooms and equipment get their own script, `New-TargetResourceMailboxes.ps1`, which rebuilds them with their capacity and booking rules. Examples use fabrikam.com as the source and contoso.com as the destination.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 2.1.1 (New-TargetMailboxes), 1.0.0 (New-TargetResourceMailboxes) |
| Released | 10 November 2026 |
| Licence | MIT |

Part 2 of my [M365 tenant-to-tenant migration toolkit](../README.md). Part 1 is [Invoke-SourceTenantInventory](../01-SourceInventory/), which produces the `04-Mailboxes.csv` and `15-ResourceMailboxes.csv` these scripts read.

## Why

Migration tools copy content. They don't decide what the destination should look like. Before any migration tool can copy a mailbox, something has to exist on the other side, with the right type, the right licence and an archive if the source has one.

Waiting for licences, or for the cutover date, holds up the copy. A shared mailbox needs no licence. So build everything as shared first, start the copy, and convert the people later. Mail lands early, and nobody can sign in to the new tenant before you're ready.

## What happens without it

- **The copy waits for licences.** Without a mailbox on the other side, no tool can start. Staging as shared means mail moves while the paperwork catches up.
- **People can sign in to the new tenant early,** see an empty mailbox, and start working in it before cutover.
- **Sign-in names land on the wrong domain,** and so do OneDrive URLs, which can't be renamed cleanly later.
- **Rooms stop taking bookings.** Recreated by hand, a room usually comes back with default settings: no capacity, no delegates, anyone can book anything forever.

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

**Tip: name the account.** With both tenants open on one PC, `Connect-ExchangeOnline` and `Connect-MgGraph` happily reuse whichever session is already there. Add `-AdminUpn admin@contoso.onmicrosoft.com` and the script disconnects any session for another account or tenant, signs in as the one you named, and prints `Signed in as X to tenant Y` before it changes anything. On the live cutover this would have saved a confusing error at the worst moment.

```powershell
.\New-TargetMailboxes.ps1 -Apply -TargetDomain contoso.onmicrosoft.com -UsageLocation GB -StageOnly -AdminUpn admin@contoso.onmicrosoft.com
```

Add `-StagingSmtpDomain contoso.com` to any `-Apply` to give staged mailboxes a primary address on a domain the destination already owns. The onmicrosoft.com address stays as a proxy, so your migration tool's mapping keeps working.

## What the plan decides

| Source | Planned as | Licence |
| --- | --- | --- |
| User mailbox | User | `-UserLicenceSku` (default `SPB`, Business Premium) |
| User mailbox listed in `-TreatAsShared` | Shared | As below |
| Shared mailbox with an archive, or near `-SharedQuotaGB` | Shared | `-SharedLicenceSku` (default `EXCHANGEENTERPRISE`, Exchange Online Plan 2) |
| Other shared mailboxes | Shared | None |
| Room or equipment mailbox | Excluded from this plan | Built by [New-TargetResourceMailboxes.ps1](#rooms-and-equipment-new-targetresourcemailboxesps1). They can't be staged as shared and converted |
| Scheduling (Bookings) mailbox | Excluded | Recreate the Bookings page in the destination. The mailbox can't be copied as a normal mailbox |

Mailboxes on the main source domain keep their alias. Mailboxes on other domains become `alias-domainname`, so `info@fabrikam.com` and `info@legacy.co.uk` don't collide. Archives at or above `-AutoExpandThresholdGB` (default 90) get auto-expanding archiving.

Open the CSV before you apply it. Two things to look for:

- **People who already exist in the destination.** Set `Include` to `False`, and map them by hand in your migration tool. Otherwise you get a duplicate mailbox that takes a licence.
- **Aliases already used in the destination.** `docs` or `marketing` often are. Change `Alias` and `TargetAddress`, e.g. to `fab-docs`.

## Rooms and equipment: New-TargetResourceMailboxes.ps1

Rooms and equipment can't be staged as shared mailboxes and converted, so they're built separately, properly, as room or equipment mailboxes. The script reads `15-ResourceMailboxes.csv` from [01](../01-SourceInventory/) (1.1.0 or later), or reads the source live with `-SourceAdminUpn`. For each one it:

1. Creates `New-Mailbox -Room` or `-Equipment` on the staging address, with the sign-in name on `-TargetDomain`
2. Copies capacity, office and the address list setting
3. Sets calendar processing: `AutomateProcessing`, `BookingWindowInDays`, `MaximumDurationInMinutes`, `AllowConflicts`, `AllowRecurringMeetings`, `AllBookInPolicy`, `AllRequestInPolicy`, plus the subject and privacy settings
4. Maps `ResourceDelegates`, `BookInPolicy` and `RequestInPolicy` through your mapping CSV. Anyone not in the mapping is left out and named in the report, never guessed
5. Tags it `CustomAttribute15 = T2T:<source address>`, so a re-run recognises it and [02b](../02b-UserMapping/) maps it

```powershell
# After 02b has built UserMapping.csv
.\New-TargetResourceMailboxes.ps1 -TargetDomain contoso.onmicrosoft.com -StagingSmtpDomain contoso.com -MappingFile .\UserMapping.csv -AdminUpn admin@contoso.onmicrosoft.com -WhatIf
.\New-TargetResourceMailboxes.ps1 -TargetDomain contoso.onmicrosoft.com -StagingSmtpDomain contoso.com -MappingFile .\UserMapping.csv -AdminUpn admin@contoso.onmicrosoft.com

# No 1.1.0 inventory? Read the source live. A snapshot CSV is saved before it signs in to the destination
.\New-TargetResourceMailboxes.ps1 -SourceAdminUpn admin@fabrikam.onmicrosoft.com -TargetDomain contoso.onmicrosoft.com -MappingFile .\UserMapping.csv -AdminUpn admin@contoso.onmicrosoft.com

# Destination already has a 'boardroom'
.\New-TargetResourceMailboxes.ps1 ... -AliasPrefix fab-
```

Then run [02b](../02b-UserMapping/) again so the rooms appear in `UserMapping.csv`, and 07 can move their old addresses at cutover.

- **Tip:** existing bookings are calendar items, so they're content. Your copy tool moves them into the new room. This script makes sure the room then behaves the same way.
- **Trap:** a delegate who has left shows as `unresolved:` in the inventory. The room is built without them. If nobody is left to approve requests, the room either accepts everything or nothing. Read the Detail column.
- **Trap:** rooms need no licence. Don't assign one by habit.
- **Trap:** Outlook's Room Finder needs room lists and Places data (building, floor). This script doesn't build those. Add them in the destination once the rooms exist.

## Output

| File | Contents |
| --- | --- |
| TargetBuildPlan.csv | The plan. Edit `Include`, `Alias`, `TargetAddress`, `Kind` and `Licence` freely |
| TargetBuild-Status-*.csv | What each run did, per mailbox, with Graph's real error text in `Note` |
| TargetBuild-Passwords-*.csv | Temporary passwords. Sensitive. Move it somewhere secure and delete the local copy |
| TargetResourceMailboxes-*.csv | What `New-TargetResourceMailboxes` did for each room and equipment mailbox: Created, Exists, Set, Skipped, Pending, Problem or WhatIf, with the booking settings and any unmapped delegates |
| ResourceMailboxes-Source-*.csv | With `-SourceAdminUpn` only: the snapshot of rooms and equipment read from the source |

## Things I hit so you don't have to

- **`Set-MgUserLicense` and the empty `-RemoveLicenses`.** The Graph SDK drops the empty array and Graph returns 400, "missing removeLicenses". The script calls the `assignLicense` API directly.
- **Passwords need their own permission.** `User.ReadWrite.All` isn't enough to set a password. The script asks for `User-PasswordProfile.ReadWrite.All` as well. If you've connected before, run `Disconnect-MgGraph` first so you get the consent prompt.
- **Set the sign-in name on shared mailboxes.** `New-Mailbox -Shared` without `-MicrosoftOnlineServicesID` puts the sign-in name on the tenant's **default** domain, not the address you gave it. In a destination that's usually the customer's real domain. It can clash with existing users, it's what people end up signing in with when you convert them, and it sets their OneDrive URL. I found out when OneDrives appeared as `name_theirdomain_com`.
- **Usage location before licence.** No usage location, no licence. The script sets it on its own, so a failed password step can't block licensing.
- **New archives aren't instant.** `ArchiveStatus` says Active straight away, but the archive itself can take a while. `Get-EXOMailboxStatistics -Archive` returning nothing means it isn't ready yet.
- **Archive copies often need EWS, and the EWS allow list changed in October 2026.** Main mail can copy fine while every archive fails with 'Unauthorized'. [Test-EwsArchiveReadiness](../03-EwsArchiveReadiness/) checks and fixes the causes in both tenants.
- **Some tools copy only email into a shared mailbox.** The one I used skipped calendars and contacts until the mailbox was a user mailbox. Check yours, and run an incremental after step 4 if needed.

## Three ways to do it

| Route | How |
| --- | --- |
| Paid tool | ShareGate and most copy tools copy into mailboxes; they don't create them. Some suites have a provisioning step: check whether it builds shared mailboxes or licensed users, and on which domain the sign-in name lands |
| Microsoft's own | Native cross-tenant mailbox migration needs **mail users**, not mailboxes, stamped with the source's ExchangeGuid and X500. Don't use this script for that route; see [option A](../copy-options/A-Microsoft-native/) |
| By hand | Exchange admin centre, Recipients, Mailboxes, Add a shared mailbox, one per person, then convert and license at cutover. Fine for five people. Remember to set the sign-in name, or it lands on the tenant's default domain. Rooms: Resources, Add a resource, then set booking options and delegates by hand from `15-ResourceMailboxes.csv` |

- **Trap: on the native route, don't license the target before the move.** A native cross-tenant move needs a mail user carrying the source ExchangeGuid (from `04-Mailboxes.csv`). Give that mail user an Exchange licence before the ExchangeGuid is stamped and Exchange builds a brand new empty mailbox instead, and the move can't land on it. Microsoft's advice is to license after the move, within the 30-day grace period. License each person once their move has completed. The Cross-Tenant User Data Migration licence is separate, and that one is needed before the move.
- **Trap:** a module version clash ("Could not load file or assembly Microsoft.Graph.Authentication") stops `-Apply` part way. Run [tools/Test-ModuleHealth.ps1](../tools/) first, and start in a fresh window.

## Notes

- `-Apply` checks that the signed-in tenant owns `-TargetDomain` before it changes anything. Run it against the wrong tenant and it stops.
- Everything supports `-WhatIf`. Use it every time you change the plan.
- Licences are only assigned from free seats. Short on seats, and the row shows `PENDING: no free <SKU>` until you buy more and re-run.

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 1.0.0 (New-TargetResourceMailboxes) | 2026-11-10 | First release. Builds room and equipment mailboxes with capacity and calendar processing, delegates and booking policies mapped, from the 01 inventory or a live read of the source |
| 2.1.1 | 2026-11-10 | Plan note for rooms and equipment points at `New-TargetResourceMailboxes.ps1` instead of "build by hand" |
| 2.1.0 | 2026-11-10 | `-AdminUpn`: disconnects a session for any other account, signs in as the one named, prints who and where before changing anything |
| 2.0.1 | 2026-11-08 | Shared mailboxes get their sign-in name set to `TargetAddress`. Without it Exchange uses the tenant's default domain. Warns when a sign-in name differs, and the passwords file shows the real one |
| 2.0.0 | 2026-11-07 | First public release. Generic parameters replace tenant-specific settings. Adds `-StagingSmtpDomain`, `-ResetPasswords`, `-TreatAsShared` and `-Exclude` |
| 1.2.2 | 2026-11-07 | Password permission added. Usage location and licence no longer depend on the password step. Resumes people converted by an interrupted run. Warnings show Graph's real error text |
| 1.2.1 | 2026-11-07 | Licences assigned through the Graph API, working round the `Set-MgUserLicense` empty-array bug |
| 1.2.0 | 2026-11-07 | `-SharedMailboxesOnly` |
| 1.1.0 | 2026-11-06 | Stage everything as shared mailboxes, convert people later |
| 1.0.0 | 2026-11-06 | First version, used on a live migration |
