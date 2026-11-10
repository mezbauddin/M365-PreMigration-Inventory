# Request-TargetOneDrives

**TL;DR:** A OneDrive doesn't exist until its owner signs in or an admin asks for one, and a migration tool can't copy into a OneDrive that isn't there. This script creates the OneDrives for everyone in your migration plan, in batches of 200, waits for them if you ask, and gives you a report with every URL. Run it as soon as people are licensed (on cutover night, straight after [02](../02-TargetMailboxes/) `-Apply`, step 13 of the [cutover runbook](../docs/Cutover-Runbook.md)), then copy after cutover.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.1.0 |
| Released | 10 November 2026 |
| Licence | MIT |

Part 4 of my [M365 tenant-to-tenant migration toolkit](../README.md). It reads the same `TargetBuildPlan.csv` that [New-TargetMailboxes](../02-TargetMailboxes/) writes.

## Why you need this

In a tenant-to-tenant move, nobody signs in to the destination until cutover. So nobody's OneDrive exists there. Point your migration tool at those users and the OneDrive copy fails or gets skipped, usually with a 'site not found' or 'OneDrive not provisioned' style error.

Creating a OneDrive isn't instant either. SharePoint queues the request, and it can take anything from a few minutes to 24 hours. Leave it until cutover night and you're waiting on a queue when you should be copying.

## What happens without it

- **The OneDrive pass fails for everyone who hasn't signed in yet.** On a tenant-to-tenant move, that's everyone.
- **You find out late.** Mail is usually copied first, so OneDrive problems show up after cutover, when people are already asking where their files are.
- **People sign in to an empty OneDrive** while you wait for provisioning, and some start saving new files there before the old ones arrive. That's messy to merge.

## How it works

1. Reads people from `TargetBuildPlan.csv` (rows with `Include` True and `Kind` User), plus anyone you pass with `-User`.
2. If Exchange Online is connected, swaps each mail address for the person's real sign-in name, and warns you where they differ.
3. Looks up which of them already have a OneDrive, by owner or by URL, and skips those.
4. Requests the rest with `Request-SPOPersonalSite`, 200 at a time (that's the cmdlet's limit).
5. With `-Wait`, checks every 5 minutes until they all exist or `-TimeoutMinutes` passes.
6. Writes `TargetOneDrives-<date>.csv` with each user, their status (Already existed, Created or Pending) and their OneDrive URL.

Re-run it as often as you like. It never requests the same OneDrive twice.

## Requirements

- Optional but recommended: run `Connect-ExchangeOnline` against the destination first, so addresses are resolved to sign-in names
- `Microsoft.Online.SharePoint.PowerShell`. In PowerShell 7 the script falls back to `-UseWindowsPowerShell` if your version of the module needs it
- SharePoint Administrator in the destination
- Users licensed with a plan that includes OneDrive (Business Premium, E3 and so on)

```powershell
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
```

## Usage

```powershell
# Everyone in the plan
.\Request-TargetOneDrives.ps1 -AdminUrl https://contoso-admin.sharepoint.com

# Wait up to two hours for them to appear
.\Request-TargetOneDrives.ps1 -AdminUrl https://contoso-admin.sharepoint.com -Wait -TimeoutMinutes 120

# Specific people, with or without a plan file
.\Request-TargetOneDrives.ps1 -AdminUrl https://contoso-admin.sharepoint.com -User kim@contoso.com, alex@contoso.com

# Preview
.\Request-TargetOneDrives.ps1 -AdminUrl https://contoso-admin.sharepoint.com -WhatIf

# Make sure the sign-in look-up runs in the destination, not whichever tenant was open last
.\Request-TargetOneDrives.ps1 -AdminUrl https://contoso-admin.sharepoint.com -AdminUpn admin@contoso.onmicrosoft.com
```

**Tip: name the account.** With both tenants open on one PC, `Connect-ExchangeOnline` and `Connect-MgGraph` happily reuse whichever session is already there. Add `-AdminUpn admin@contoso.onmicrosoft.com` and the script disconnects any session for another account or tenant, connects Exchange Online as the one you named (for the address to sign-in look-up), and prints `Signed in as X to tenant Y`. On the live cutover this would have saved a confusing error at the worst moment.

## Tips and traps

- **Tip:** run it on the night with `-Wait -TimeoutMinutes 60` if you can stay up, so the morning report shows Created rather than Pending.
- **Tip:** re-run it at T+1 before you build the OneDrive copy list. Anyone still Pending gets requested again; anyone Created is skipped.
- **Trap:** run it before the sign-in names switch and the OneDrive gets an onmicrosoft.com URL that SharePoint then renames. That works, but your copy list must use the final URL, so build the list after cutover.
- **Trap:** OneDrive sharing links and "Shared with me" don't come across with the files. See [What doesn't move](../docs/What-Doesnt-Move.md).

## Three ways to do it

| Route | How |
| --- | --- |
| Paid tool | ShareGate can't copy into a OneDrive that doesn't exist, and doesn't create them. Run this first whatever tool you use |
| Microsoft's own | `Request-SPOPersonalSite -UserEmails` by hand, 200 at a time, then `Get-SPOSite -IncludePersonalSite $true` to see which exist. Exactly what this script does, without the report. Native cross-tenant OneDrive migration creates the destination OneDrive itself |
| By hand | Have each person sign in to office.com once. Their OneDrive appears within minutes. Not possible before cutover, when nobody can sign in yet |

## Lessons learnt

- **Provision early, copy late.** The OneDrive URL is built from the user's sign-in name. If you stage people on `onmicrosoft.com` and switch them to the real domain at cutover, SharePoint renames the OneDrive to match. Create it early so it's ready, and run the copy after the rename so your mapping points at the final address.
- **The OneDrive follows the sign-in name, not the mail address.** I requested OneDrives using people's staging addresses and two came back as `name_customerdomain_com`. Their sign-in names had been put on the destination's default domain when the mailboxes were created. Check sign-in names before you provision. This script does when Exchange is connected, and [New-TargetMailboxes](../02-TargetMailboxes/) 2.0.1 now sets them properly.
- **A new OneDrive has no owner for a while.** `Get-SPOSite` shows it with a blank Owner at first, so the script matches on the URL too.
- **Licence first.** No OneDrive licence, no OneDrive. Run [New-TargetMailboxes](../02-TargetMailboxes/) to license people before this.
- **Don't trust 'requested' as 'done'.** The cmdlet returns straight away whatever happens. Check that the site exists, which is what `-Wait` and the report do.
- **Mind people who already exist.** Anyone who was already in the destination tenant may have a OneDrive full of their own files. The report marks them 'Already existed'. Decide whether to merge into it before you copy.

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| README | 2026-11-10 | Tips and traps, runbook timing. Script unchanged at 1.1.0 |
| 1.1.0 | 2026-11-10 | `-AdminUpn`: connects Exchange Online as the named account (disconnecting any other) so sign-in names are looked up in the right tenant |
| 1.0.1 | 2026-11-08 | Resolves mail addresses to sign-in names through Exchange Online. Matches new OneDrives by URL as well as owner |
| 1.0.0 | 2026-11-08 | First release |
