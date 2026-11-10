# 08 Device rejoin

**TL;DR:** A Windows laptop joined to the source tenant's Entra ID has to leave it and join the destination, and every app on it has to sign in again. `Get-DeviceInventory.ps1` tells you which devices need it, who uses them and whether Intune manages them. `Export-BitLockerKeys.ps1` saves every BitLocker recovery key before the source can lose them. The printable checklist below takes one device from "still in the old tenant" to "working in the new one" in 45 to 60 minutes without losing anyone's files. Laptops joined to an on-premises Active Directory are unaffected by the tenant move. Examples use fabrikam.com as the source and contoso.com as the destination.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.0.0 (Get-DeviceInventory, Export-BitLockerKeys) |
| Released | 10 November 2026 |
| Licence | MIT |

Part 8 of my [M365 tenant-to-tenant migration toolkit](../README.md). Run the scripts before cutover. Rejoin devices from the first working day after cutover ([cutover runbook](../docs/Cutover-Runbook.md), T+1).

## Why you need it

A Windows device can be Entra joined to one tenant only. After cutover the person's source account is blocked, so the device they log on to every morning still trusts a tenant that no longer knows them. Cached credentials keep them working for a day or two, then it falls apart: Windows Hello stops, OneDrive stops syncing, Outlook prompts forever, and Intune policy and apps from the source keep applying.

## What happens without it

- **People log on with cached credentials** and think all is well until the cache expires, usually on a Monday.
- **OneDrive keeps syncing to the old tenant**, or to nowhere, and files saved after cutover go to the wrong place.
- **The device stays managed by the source Intune**, so the source can still wipe it, and the destination can't manage it.
- **A local profile is lost** when the device is reset or rejoined without a backup, including desktop files that were never in OneDrive.
- **A BitLocker recovery prompt becomes a dead laptop.** The recovery key lives in the source tenant. Close the source without exporting it and nobody can unlock that disk again.

## Which devices are affected

| Join type (`trustType`) | What it is | What to do |
| --- | --- | --- |
| Entra joined (`AzureAd`) | Signed in with a work account, no on-prem domain | Rejoin with the checklist below |
| Registered (`Workplace`) | Personal PC or phone with a work account added | Remove the work account, add the new one. No rejoin |
| Hybrid joined (`ServerAd`) | Joined to on-prem Active Directory, synced to Entra ID | **Unaffected by the tenant move.** The domain join stays. Its Entra registration follows whichever tenant Entra Connect syncs to; only the apps need signing in again |
| On-prem AD only | Joined to Active Directory, not in Entra ID | **Unaffected.** Sign the apps in again |

## The scripts

### Get-DeviceInventory.ps1

Read-only. Reads every Entra ID device through Microsoft Graph with its join type, OS, registered owner and users, approximate last sign-in and, unless `-SkipIntune`, the matching Intune record (compliance, last sync, serial number, model). Each row gets an Action: Rejoin, Re-register, On-prem AD: unaffected, or Stale. It also tells you roughly how many hours of hands-on time the rejoins need.

```powershell
.\Get-DeviceInventory.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com
# After the rejoins: the same against the destination, to check every device arrived
.\Get-DeviceInventory.ps1 -TenantId <dest-guid> -AdminUpn admin@contoso.onmicrosoft.com
```

Needs `Microsoft.Graph.Authentication`, Global Reader, and the Graph scopes `Device.Read.All`, `User.Read.All` and `DeviceManagementManagedDevices.Read.All` (that last one only for Intune). Writes `DeviceInventory-<tenant>-<date>.csv`.

### Export-BitLockerKeys.ps1

Read-only in the tenant. Lists every recovery key in `informationProtection/bitlocker/recoveryKeys`, reads each key value, matches it to the device name, owner and OS, and writes one CSV into a folder that only you can read. It refuses to write into a OneDrive, SharePoint or other synced folder, or inside a Git repository, unless you pass `-Force`. It also lists Entra joined Windows devices that have no key in Entra ID at all.

```powershell
.\Export-BitLockerKeys.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -OutputFolder D:\Secure\BitLocker
```

Needs `Microsoft.Graph.Authentication`, the Graph scopes `BitLockerKey.Read.All` and `Device.Read.All`, and a role that can read keys: Cloud Device Administrator, Helpdesk Administrator, Intune Administrator, Security Administrator, Security Reader or Global Reader.

**Trap:** that CSV opens every disk in the company. Move it into your password manager or an encrypted store the business controls, and delete the local copy. Every key read is audited in Entra ID, so tell the security team before you run it.

### Copy-ProfileData.ps1

Runs **on the laptop**, elevated. `-Prepare` before it leaves the old tenant saves the BitLocker key, join status, printers and profile list to `C:\MigrationBackup`. After the person has signed in once with their new account, the default mode copies Desktop, Documents, Downloads, Pictures, Videos, Music, Favorites, the old OneDrive folder(s), Outlook signatures and PST files, Sticky Notes and Edge/Chrome bookmarks from the old profile into the new one. The old profile is never changed, so it's your fallback.

```powershell
.\Copy-ProfileData.ps1 -Prepare                          # before leaving the old tenant
.\Copy-ProfileData.ps1 -OldProfile jsmith -WhatIf       # after first sign-in with the new account
.\Copy-ProfileData.ps1 -OldProfile jsmith
```

- **Tip:** it skips cloud-only OneDrive placeholders on purpose. Once the old account is blocked they can't download, and they arrive with the OneDrive copy anyway.
- **Trap:** saved browser passwords are encrypted to the old account and can't be copied. Export them (Edge: Settings > Passwords > Export) while the person can still sign in to the old profile.
- **Three ways:** ForensiT User Profile Wizard (paid, re-uses the old profile in place), this script (free, copies into the new profile), or copying the folders by hand in File Explorer as an admin.

## Per-device checklist

Print one per device. Do it with the person beside you: they know what's on the desktop and which browser they use.

**Device:** ____________________ **Person:** ____________________ **New sign-in:** ____________________@contoso.com

**Date:** __________ **Start:** __________ **Finish:** __________ **Done by:** ____________________

### Before you start (about 15 minutes)

- [ ] Run `dsregcmd /status` in a command prompt. Under Device State, `AzureAdJoined : YES` and `DomainJoined : NO` means Entra joined: carry on. `DomainJoined : YES` means on-prem AD: stop, the device stays as it is, sign the apps in again only. Note `TenantName` (should be the source)
- [ ] BitLocker key for this device is in the export from `Export-BitLockerKeys.ps1`. If not, run `manage-bde -protectors -get C:` and write down the numerical password now
- [ ] **Create a local admin account and test it.** Settings, Accounts, Other users, Add account, "I don't have this person's sign-in information", "Add a user without a Microsoft account". Make it an administrator. Sign out, sign in as it, sign out. Without this there's no way back in after you disconnect
- [ ] OneDrive: the cloud icon shows "Up to date". Known folders (Desktop, Documents, Pictures) are backed up: OneDrive settings, Sync and back up, Manage back up
- [ ] Copy anything outside OneDrive to a USB drive or a folder the new account can reach:
  - [ ] Browser data: sign in to Edge or Chrome sync with a personal account, or export bookmarks to a file. Saved passwords: export them, or ask the person to use a password manager
  - [ ] Outlook signatures: `%APPDATA%\Microsoft\Signatures`
  - [ ] PST files: search the drive for `*.pst`, usually under Documents\Outlook Files
  - [ ] Sticky Notes, Downloads, anything saved to `C:\` or a second drive
- [ ] Note mapped drives (`net use`) and printers (Settings, Bluetooth and devices, Printers and scanners)

### Leave the source (about 5 minutes)

- [ ] Settings, Accounts, Access work or school, select the source account, **Disconnect**, confirm. Or, from an elevated prompt, `dsregcmd /leave`
- [ ] Enter the local admin account when asked
- [ ] **Reboot.** Sign in as the local admin
- [ ] `dsregcmd /status` shows `AzureAdJoined : NO`

### Join the destination (about 10 minutes)

- [ ] Settings, Accounts, Access work or school, **Connect**, then **Join this device to Microsoft Entra ID** (the link below the email box, not the box itself)
- [ ] Sign in with the person's new sign-in name and temporary password. Confirm it's the right organisation, then Join
- [ ] Reboot. Sign in as the person (Other user, `name@contoso.com`). They set their password and register MFA in the destination
- [ ] `dsregcmd /status` shows `AzureAdJoined : YES` and `TenantName` is the destination

### Intune and Windows Hello (about 10 minutes)

- [ ] Intune enrolment: automatic if MDM user scope is set in the destination. Settings, Accounts, Access work or school, the account, Info, shows the management server. The device appears in the Intune admin centre within about 15 minutes
- [ ] Windows Hello: the old PIN and face or fingerprint belonged to the source account. Set up a new PIN when prompted, or Settings, Accounts, Sign-in options. If Hello misbehaves, remove the PIN there and set it again

### Profile and apps (about 15 to 20 minutes)

- [ ] The person has a new, empty Windows profile. The old one is still in `C:\Users\<old name>`. Copy across what OneDrive didn't hold, from the backup or the old profile folder (local admin can read it)
- [ ] OneDrive: sign in with the new account, choose the folder, let it sync. Turn on known folder backup again. Don't point it at the old OneDrive folder
- [ ] Outlook: new profile, signed in with the new account. Shared mailboxes appear on their own (automapping). Put signatures back. Open any PST files the person still needs (File, Open and export)
- [ ] Teams: sign out of the old account, sign in with the new one. Clear the old account from the account list
- [ ] Office apps: File, Account, sign out of the old account, sign in with the new. Check activation
- [ ] Browser: sign in to the work profile with the new account. Re-import bookmarks
- [ ] Printers and mapped drives: add back from the notes above
- [ ] Phone (if they have one): remove the work account from Outlook, Teams and Authenticator, add the new one

### Finish

- [ ] The person signs in, opens a file from OneDrive, sends an email and joins a Teams call
- [ ] Remove the local admin account, or disable it, as your policy says. Write down which
- [ ] Old profile folder: leave it for a week, then delete it once the person confirms nothing is missing
- [ ] Note the time taken. It tells you how many devices support can do on day one

## Timings

Plan **45 to 60 minutes per device** with the person present, after the first two or three. The first one always takes longer: time it before the night. A large unsynced Desktop or many PST files can double it. Two people working in parallel can usually do six to eight devices in a morning. `Get-DeviceInventory.ps1` prints the total hands-on time for the devices it finds.

## Three ways to do it

| Route | How |
| --- | --- |
| Paid tool | Profile migration tools such as ForensiT User Profile Wizard re-point the existing Windows profile to the new account, so the desktop, settings and files stay where they are and the person sees no new profile. Paid editions can join the device to the destination Entra ID and run unattended. Some migration suites have a device agent that does the same. Test one laptop, end to end, before you buy |
| Microsoft's own | Windows Autopilot: export the device hashes, delete the devices from the source Autopilot, import them into the destination, then reset each device. Clean, and Intune builds it the way the destination wants, but the device is wiped, so the profile backup in the checklist matters even more. Free with Intune |
| By hand | The checklist above, per device, from Settings. Free, no wipe, and you see every problem as it happens. Fine for a small office |

## Tips and traps

- **Trap: no local admin, no way back.** Disconnecting from Entra ID removes the only admin account many devices have. Create a local admin, and test it, before you disconnect.
- **Trap: a new profile.** Joining the destination creates a new Windows profile for the new account. The old one stays on disk but the person won't see it. Copy what OneDrive didn't hold, or use a profile tool.
- **Trap: cached credentials hide the problem.** Blocking source sign-in doesn't stop a device logging on with cached credentials until the cache expires. Don't treat "they can still log on" as success.
- **Trap: BitLocker keys stay in the source.** Export them with `Export-BitLockerKeys.ps1` before any device leaves. After the rejoin, check the destination has a new key for the device (Entra admin centre, Devices, the device, BitLocker keys). If not, run `BackupToAAD-BitLockerKeyProtector -MountPoint C: -KeyProtectorId <id>` in an elevated PowerShell window (the ID comes from `manage-bde -protectors -get C:`), or let an Intune BitLocker policy do it.
- **Trap: the device is still in the source Intune.** Retire or delete it there once it's rejoined, so the source can't wipe it later, and so the old record doesn't confuse compliance reports.
- **Trap: Conditional Access in the destination.** A "require compliant device" policy blocks the person until Intune has finished. Allow time, or exclude them for an hour.
- **Tip:** the OneDrive URL changes with the sign-in name. Sign OneDrive in fresh rather than trying to re-point the old sync relationship.
- **Tip:** do one device end to end before the night, and time it. That number sets how many people support can handle on day one.
- **Tip:** hybrid joined and on-prem AD devices don't need any of this. Don't let them fill the day-one queue.

## Output

| File | Contents |
| --- | --- |
| `DeviceInventory-<tenant>-<date>.csv` | Every Entra ID device: join type, OS, owner, users, last sign-in, Intune state, serial number, model, Action |
| `BitLockerKeys-<tenant>-<date>\BitLockerKeys-<tenant>-<date>.csv` | Device name, owner, volume, key ID and **recovery key**. Most sensitive file in the toolkit |

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 1.1.0 | 2026-11-10 | Added `Copy-ProfileData.ps1`: pre-move snapshot and old-to-new profile copy on the laptop |
| 1.0.0 | 2026-11-10 | First release of the scripts: `Get-DeviceInventory.ps1` and `Export-BitLockerKeys.ps1`. Printable per-device checklist, timings, three ways. Replaces the planning page |
