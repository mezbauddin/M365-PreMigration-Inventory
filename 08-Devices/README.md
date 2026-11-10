# 08 Device rejoin

**TL;DR:** A Windows laptop joined to the source tenant's Entra ID has to leave it and join the destination, and every app on it has to sign in again. `Get-DeviceInventory.ps1` tells you which devices need it, who uses them and whether Intune manages them. `Export-BitLockerKeys.ps1` saves every BitLocker recovery key before the source can lose them. `Move-TenantDevices.ps1` is one menu for the rest: build the bulk enrolment package, copy Intune configuration, move Autopilot hashes, then move each laptop in stages and hand the person's existing profile to their new account. The printable checklist below takes one device from "still in the old tenant" to "working in the new one" in 45 to 60 minutes without losing anyone's files. Laptops joined to an on-premises Active Directory are unaffected by the tenant move. Examples use fabrikam.com as the source and contoso.com as the destination.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.2.1 (Move-TenantDevices 1.0.1, Copy-ProfileData 1.0.0, Get-DeviceInventory 1.0.0, Export-BitLockerKeys 1.0.0) |
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

### Move-TenantDevices.ps1 (start here)

One interactive script for the whole device side. Run it with no parameters and pick from the menu. Options 1 to 3 and 7 run on your admin PC; 4 to 6 run on each laptop as administrator. Every step that changes a tenant has `-WhatIf`, and nothing is ever wiped, retired or reset.

#### Where to run what

| Options | Run on | Why |
|---|---|---|
| 1, 2, 3 > move, 7 | Any admin PC. Your own laptop is fine | They sign in to tenants. Option 1 installs the designer tools it needs; option 7 removes them again |
| 3 > capture, 4, 5, 6 | Only the device being moved, as administrator | They read and change that machine: its tenant join, local accounts and profiles |
| 4 | **Never on your own PC** | Option 4 takes the machine it runs on out of its tenant. On your admin laptop that means you, not the user, lose your work account sign-in |

Option 4 has a guard for exactly that mistake. Pass `-SourceTenantId` and stage 1 stops, before changing anything, if `dsregcmd /status` shows the machine joined to any other tenant. Leave it out and the script shows the joined tenant's name and ID and makes you type the name exactly before it leaves; `LEAVE` alone won't do. Either way it also refuses to leave unless the local admin safety-net account exists, is enabled and is in the local Administrators group.

| Option | Where | What it does |
|---|---|---|
| 1 Build the enrolment package | Admin PC | Signs in to the destination, gets a bulk enrolment token and builds a `.ppkg` that joins laptops silently. One per tenant, valid up to 180 days. Asks nothing: tenant and account come from the sign-in. `-RenamePattern`, `-PackageName`, `-TokenDays` to override; `-AllowWizard` to fall back to the designer's wizard |
| 2 Copy Intune configuration | Admin PC | Exports device configuration, settings catalog, compliance, scripts, remediations, Autopilot profiles and app metadata from the source; imports into the destination; maps group assignments by name; skips anything with the same name already there |
| 3 Autopilot | Laptop then admin PC | Capture the hardware hash on the laptop; register captured hashes in the destination and remove them from the source |
| 4 Move this laptop | Laptop | Staged, with the stage saved in `C:\MigrationBackup\state.json`. Stage 1: pre-flight, BitLocker key, printers, Autopilot hash, local admin, leave the old tenant. Stage 2 (after a restart as the local admin): install the package, join. Stage 3: the person signs in once |
| 5 Hand the old profile over | Laptop | Points the person's new account at their old profile folder, fixes ownership and permissions (files and registry hives), parks the empty new profile. Verifies, and rolls back on any failure |
| 6 Copy profile data instead | Laptop | Fallback: `Copy-ProfileData.ps1` |
| 7 Clean up this admin PC | Admin PC | Removes what option 1 installed or left behind: private designer copies, temp files, the sign-in window setting, build logs and, once you confirm, Windows Configuration Designer and the ADK. Keeps your `.ppkg` files |

```powershell
# Admin PC, once per destination
.\Move-TenantDevices.ps1 -Action BuildPackage -RenamePattern 'TG-%SERIAL%'   # 1: sign in once, nothing else to type
.\Move-TenantDevices.ps1                      # then 2 (with -WhatIf first), 3 > move, and 7 when you are finished

# Each laptop being moved (never your own PC), as administrator, re-run after each restart until it says Done
.\Move-TenantDevices.ps1 -Action MoveDevice -SourceTenantId <source-guid> -PackagePath C:\MigrationBackup\contoso-join.ppkg
```

**What the person sees:** their laptop restarts twice, they sign in once with their new email and password, you run option 5, and their next sign-in has their desktop, files, settings and apps exactly as they were. OneDrive, Outlook and Teams ask for the new sign-in once. Saved browser passwords and Windows Hello are tied to the old account and need setting up again.

- **Tip:** do one willing person first, end to end, before the rest.
- **Tip:** option 2 finds nothing in a tenant that never used Intune. That's fine; the laptops still enrol in the destination's Intune automatically on join (MDM enrolment scope).
- **Trap:** an ordinary sign-in token is not a bulk token. A package built from one won't join anything; the bulk token comes from the `bulkaadjtoken` endpoint, the same one Windows Configuration Designer's "Get Bulk Token" button calls. Option 1 uses it.
- **Trap:** the bulk token request in option 1 uses the same undocumented sign-in flow as Windows Configuration Designer. If Microsoft changes it, option 1 stops with the exact error and opens nothing; add `-AllowWizard` to fall back to the designer's wizard.
- **Trap:** the `.ppkg` joins any device to the tenant. Keep it off shared drives and delete it when the move is done.
- **Trap:** run stages 2 and 5 from the local admin, never from the person's account. A loaded profile can't be handed over.
- **Trap:** option 4 on your own admin PC removes it from its tenant. Always pass `-SourceTenantId`, and if the script stops with "joined to ... not the source tenant", you're on the wrong machine.
- **Three ways:** ForensiT User Profile Wizard (paid, polished, does the handover for you), this script (free, same handover technique), or by hand (leave, join, copy the files with Explorer and lose the settings).

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

- **The current ADK designer crashes from the command line.** ICD.exe in ADK 10.1.26100 and later is missing `Microsoft.Diagnostics.Tracing.EventSource.dll` (FileNotFoundException). Option 1 runs a private copy with that file added from its NuGet package (1.1.28).
- **Windows PowerShell 5.1 + `2>&1` + `$ErrorActionPreference = 'Stop'` = a crash on any native stderr.** icacls always prints "Access is denied" for the legacy junctions in a profile (`Application Data`, `Cookies`). Relax the preference around native calls and verify the result instead.
- **"Profile still loaded" after sign-out is normal.** Switch user, a lingering session or a background process keeps the hive open. A restart, then signing in as the local admin only, clears it every time.
- **Copy the person's OneDrive before you move their laptop.** With Files On-Demand most OneDrive files are cloud-only placeholders: they look present (full size in Explorer) but no data is on the disk, so nothing on the laptop can copy them once the old account is blocked. Only the tenant-to-tenant OneDrive copy brings them. Do it first, and the files are waiting at first sign-in.
- **Profile handover can end in a black screen.** In the field it mapped the profile correctly but the first sign-in hung on the per-user setup screens (UserOOBEBroker, Windows Hello) with a vendor agent crashing alongside. Copying into the fresh profile worked first time. Copy is now the default.
- **Look for duplicate destination accounts before copying.** An unlicensed `@tenant.onmicrosoft.com` twin of each person (from a set-up step run twice) makes tools guess the wrong OneDrive. Group users by display name, compare licences and addresses, prove which one the laptop uses (an Entra SID `S-1-12-1-a-b-c-d` is the account's object ID), then remove the twin.
- **Same email, different tenant: clear the cached sign-ins.** The handed-over profile still holds the old tenant's tokens (Microsoft.AAD.BrokerPlugin, Office `Common\Identity`, Teams cache). Left alone, Office and Teams loop on "Something went wrong". Option 5 renames or exports them. OneDrive still needs Unlink this PC and a new sign-in.
- **Entra profiles don't always show as `AzureAD\name`.** On current Windows they can show a tenant prefix (`CONTOSO\name`). Match them by SID instead: Entra user SIDs start `S-1-12-1-`.
- **Over Quick Assist, typing is where mistakes happen.** Every choice in options 4 and 5 is a numbered list; irreversible steps need an explicit 1 (Enter alone does nothing).
- **32-bit PowerShell can't see dsregcmd.** Some "Run as administrator" shortcuts open Windows PowerShell (x86). There `dsregcmd` is not found and BitLocker and registry calls are redirected. The script now relaunches itself in 64-bit; by hand, run `C:/Windows/Sysnative/WindowsPowerShell/v1.0/powershell.exe` first.
- **The Windows sign-in broker can't get this token.** The first-party apps allowed to request a device registration token have no broker redirect URI, so WAM fails with AADSTS50011. Option 1 uses a sign-in window instead.
- **Each bulk token request creates a `package_<guid>` account** in the destination, exactly as the designer does. Failed or repeated runs leave extra ones; delete the unused ones, never the one behind the package you are using (deleting it kills that package).
- **Device code is often blocked.** Many tenants block device code flow with Conditional Access. Option 1 signs in with a Microsoft sign-in window or the browser instead, so MFA and Conditional Access apply as normal. Use `-UseDeviceCode` only if you know it is allowed.
- **Trap: no local admin, no way back.** Disconnecting from Entra ID removes the only admin account many devices have. Create a local admin, and test it, before you disconnect.
- **Trap: a new profile.** Joining the destination creates a new Windows profile for the new account. The old one stays on disk but the person won't see it. Copy what OneDrive didn't hold, or use a profile tool.
- **Trap: cached credentials hide the problem.** Blocking source sign-in doesn't stop a device logging on with cached credentials until the cache expires. Don't treat "they can still log on" as success.
- **Trap: BitLocker keys stay in the source.** Export them with `Export-BitLockerKeys.ps1` before any device leaves. After the rejoin, check the destination has a new key for the device (Entra admin centre, Devices, the device, BitLocker keys). If not, run `BackupToAAD-BitLockerKeyProtector -MountPoint C: -KeyProtectorId <id>` in an elevated PowerShell window (the ID comes from `manage-bde -protectors -get C:`), or let an Intune BitLocker policy do it.
- **Trap: the device is still in the source Intune.** Delete its old record there once it's rejoined, so the old record doesn't confuse compliance reports. Delete only: never wipe or retire, because those actions can still reach a device that hasn't checked in for a while ([10 Decommission](../10-Decommission/)).
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
| 1.5.0 | 2026-11-10 | Move-TenantDevices 1.3.0: copy (option 6) is the recommended route, handover (option 5) warns first. Copy-ProfileData 1.0.1 counts cloud-only OneDrive placeholders separately. Order: copy each person's OneDrive before moving their laptop. |
| 1.4.4 | 2026-11-10 | Move-TenantDevices 1.2.4 and 1.2.5: handover progress messages; clearing old sign-ins no longer stops on a missing Office registry key. First real laptop moved end to end. |
| 1.4.3 | 2026-11-10 | Move-TenantDevices 1.2.3: option 5 no longer aborts on icacls "Access is denied" for legacy junctions (Windows PowerShell stderr); offers a restart if a profile is still open; `-Action Update` (menu 8) pulls the latest version. |
| 1.4.2 | 2026-11-10 | Move-TenantDevices 1.2.2: option 5 clears the old tenant's cached sign-ins (sign-in broker, Office identities, Teams cache); renamed or exported first, so reversible. |
| 1.4.1 | 2026-11-10 | Move-TenantDevices 1.2.1: stage 2 finds the `.ppkg` if it was copied in after stage 1; the generated local admin password is short and easy to type (`Move1234!`) because the account is temporary. |
| 1.4.0 | 2026-11-10 | Move-TenantDevices 1.2.0: pick by number instead of typing (profiles, package, confirmations); can generate the local admin password; package auto-found; Entra profiles detected by SID so option 5 works when Windows shows a tenant prefix instead of `AzureAD\`. |
| 1.3.1 | 2026-11-10 | Move-TenantDevices 1.1.1: relaunches itself in 64-bit PowerShell when started from 32-bit (x86), where `dsregcmd` is not found. |
| 1.3.0 | 2026-11-10 | Move-TenantDevices 1.1.0: option 7 cleans up the admin PC (designer copies, temp files, sign-in setting, build logs, and on confirmation Windows Configuration Designer and the ADK). Keeps .ppkg files. |
| 1.2.10 | 2026-11-10 | Move-TenantDevices 1.0.10: package build uses absolute paths; the ADK designer (10.1.26100 and later ships without `Microsoft.Diagnostics.Tracing.EventSource.dll`) runs from a private patched copy. |
| 1.2.9 | 2026-11-10 | Move-TenantDevices 1.0.9: option 1 sign-in is plain PowerShell (OAuth code flow in a sign-in window, then the browser with PKCE). No dependency on MSAL files from other modules. |
| 1.2.8 | 2026-11-10 | Move-TenantDevices 1.0.8: option 1 asks nothing. Tenant, account and package name come from the sign-in; `-PackageName`, `-TokenDays`, `-RenamePattern` override. Windows broker dropped (AADSTS50011 for the device registration clients). |
| 1.2.7 | 2026-11-10 | Move-TenantDevices 1.0.7: option 1 sign-in runs in a clean child PowerShell, so it works when Microsoft Graph is already loaded. |
| 1.2.6 | 2026-11-10 | Move-TenantDevices 1.0.6: option 1 is no-touch. Builds from the command line (ADK designer, or the Store designer copied out of its sandbox), installs the ADK designer silently if needed, shows the exact error and never opens the wizard unless `-AllowWizard`. |
| 1.2.5 | 2026-11-10 | Move-TenantDevices 1.0.5: option 1 signs in interactively (Windows sign-in broker, then browser) so MFA and Conditional Access apply. Device code only with `-UseDeviceCode`. |
| 1.2.4 | 2026-11-10 | Move-TenantDevices 1.0.4: option 1 builds the package with no clicks (real bulk token, silent ICD build) |
| 1.2.3 | 2026-11-10 | Move-TenantDevices 1.0.3: guided build finds the new .ppkg itself |
| 1.2.2 | 2026-11-10 | Move-TenantDevices 1.0.2: option 1 finds or installs Windows Configuration Designer (Store or ADK) |
| 1.2.1 | 2026-11-10 | `Move-TenantDevices.ps1` 1.0.1: option 4 safety guard (`-SourceTenantId`, typed tenant name, local admin must exist and be enabled). "Where to run what" table |
| 1.2.0 | 2026-11-10 | Added `Move-TenantDevices.ps1`: one menu for package, Intune copy, Autopilot, staged device move and profile handover |
| 1.1.0 | 2026-11-10 | Added `Copy-ProfileData.ps1`: pre-move snapshot and old-to-new profile copy on the laptop |
| 1.0.0 | 2026-11-10 | First release of the scripts: `Get-DeviceInventory.ps1` and `Export-BitLockerKeys.ps1`. Printable per-device checklist, timings, three ways. Replaces the planning page |
