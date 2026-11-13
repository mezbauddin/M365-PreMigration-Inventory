# 08 Device rejoin

**TL;DR:** A Windows laptop joined to the source tenant's Entra ID has to leave it and join the destination, and the person's files have to come with them. `Move-TenantDevices.ps1` option 9 does it **in one go**: one remote session while the person is still signed in, about 10 minutes of your time. It saves the BitLocker key, creates a safety-net local admin, leaves the source, joins the destination with a bulk enrolment package, and restarts. When the person signs in with their new email, a one-off task copies their files into the new profile and tells them when it's done. It also moves stand-alone PCs that were never joined. `Get-DeviceInventory.ps1` finds the devices and `Export-BitLockerKeys.ps1` saves recovery keys first. Copy each person's OneDrive **before** their laptop moves. Laptops joined to an on-premises Active Directory are unaffected. Examples use fabrikam.com as the source and contoso.com as the destination.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 2.0.0 (Move-TenantDevices 1.4.6, Copy-ProfileData 1.0.1, Get-DeviceInventory 1.0.0, Export-BitLockerKeys 1.0.0) |
| Released | 13 November 2026 |
| Licence | MIT |

Part 8 of my [M365 tenant-to-tenant migration toolkit](../README.md). Run the inventory and key export before cutover. Move devices from the first working day after cutover ([cutover runbook](../docs/Cutover-Runbook.md), T+1), each one after that person's OneDrive copy has finished.

**Field result:** eight Windows computers moved for a small firm, seven Entra joined laptops and one stand-alone PC, with nobody in the office and all work over Quick Assist. After the first pilot, each took about 10 minutes of hands-on time; people signed in with their new account and found their files already there.

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

One interactive script for the whole device side. Run it with no parameters for the menu, or pass `-Action`. Every step that changes a tenant has `-WhatIf`, and nothing is ever wiped, retired or reset.

| Option | Where | What it does |
|---|---|---|
| **9 Move this laptop in one go** | The laptop, Terminal (Admin) | **Recommended.** Run once while the person is still signed in. Finds the package (next to the script, in `C:\MigrationBackup`, in their Downloads or Desktop, or at a download link), saves the BitLocker key, printers and profile list, creates or resets the safety-net local admin and **tests its password**, leaves the source, joins the destination, registers a one-off SYSTEM task, restarts. At the person's first sign-in with the new account the task copies their old profile's files into the new one (option 6 logic), backs the BitLocker key up to the destination, shows a message and removes itself. Stand-alone PCs (local Windows account, never joined) skip the leave |
| 1 Build the enrolment package | Admin PC | Signs in to the destination (interactive, MFA and Conditional Access apply), gets a bulk enrolment token and builds a `.ppkg` that joins laptops silently. One per tenant, valid up to 180 days. Asks nothing: tenant and account come from the sign-in. `-RenamePattern`, `-PackageName`, `-TokenDays` to override |
| 2 Copy Intune configuration | Admin PC | Exports device configuration, settings catalog, compliance, scripts, remediations, Autopilot profiles and app metadata from the source; imports into the destination; maps group assignments by name |
| 3 Autopilot | Laptop then admin PC | Capture the hardware hash on the laptop; register captured hashes in the destination and remove them from the source |
| 4 Move step by step | Laptop | The staged route: stage 1 leaves, restart, stage 2 as the local admin joins, the person signs in, then option 5 or 6. Also the retry if option 9's join didn't finish |
| 5 Hand the old profile over | Laptop | Advanced. Points the new account at the old profile folder. Asks before it starts: in the field it produced a black screen at first sign-in (see Tips and traps). Use 6 |
| 6 Copy profile data | Laptop | `Copy-ProfileData.ps1` into the new profile. Option 9 runs this for you |
| 7 Clean up this admin PC | Admin PC | Removes what option 1 installed or left behind. Keeps your `.ppkg` files |
| 8 Update | Anywhere | Downloads the latest version of the script |

```powershell
# Admin PC, once per destination
.\Move-TenantDevices.ps1 -Action BuildPackage -RenamePattern 'CT-%SERIAL%'

# Each computer, in the person's own session, Terminal (Admin). Send them these three lines in Teams first.
mkdir C:/MigrationBackup -Force
iwr https://raw.githubusercontent.com/mezbauddin/M365_Tenant_to_Tenant_Migration_Toolkit/main/08-Devices/Move-TenantDevices.ps1 -OutFile C:/MigrationBackup/Move-TenantDevices.ps1
powershell -ExecutionPolicy Bypass -File C:/MigrationBackup/Move-TenantDevices.ps1 -Action MoveInOneGo -SourceTenantId <source-guid> -LocalAdminPassword <one password for every laptop>
# Stand-alone PC (local Windows account, not joined anywhere): leave out -SourceTenantId
```

Then: Enter (their profile is first), `1` (password noted), `1` (move). It restarts by itself.

**What the person sees:** one restart. On the sign-in screen they choose **Other user**, sign in with their new email, set a new PIN, and within a few minutes a message says their files are in place. Then OneDrive, Outlook and Teams ask for the new sign-in once. Saved browser passwords and Windows Hello are tied to the old account and need setting up again.

**Getting the package onto the laptop:** option 9 looks next to the script, in `C:\MigrationBackup`, in the person's Downloads and Desktop, then at `-PackageUrl` (a private, expiring link), and asks you to paste a link if it still has none. Never publish the package: it joins any device to the tenant until its `package_<guid>` account is deleted. Deleting that account kills the package everywhere, history included; joined devices are unaffected.

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

## Per-device checklist (option 9)

**Device:** ____________________ **Person:** ____________________ **New sign-in:** ____________________@contoso.com

**Date:** __________ **Start:** __________ **Finish:** __________ **Done by:** ____________________

### Before the call (your PC, 2 minutes)

- [ ] That person's **OneDrive copy has finished** (and errors retried). With Files On-Demand most of their files exist only in the cloud; nothing on the laptop can copy them once the old account is blocked
- [ ] They know their **new password** (reset it with "change at next sign-in" if not)
- [ ] One account per person in the destination: no unlicensed `@tenant.onmicrosoft.com` twin from a set-up step run twice
- [ ] Teams them the three lines above (Teams keeps them intact; email filters and link rewriting can break them)

### On the call (about 10 minutes)

- [ ] Quick Assist (`Win + Ctrl + Q`), the person signed in as themselves
- [ ] In their normal PowerShell, `net localgroup administrators`: note who can approve the admin prompt. Usually the account that joined the laptop, not your global admin
- [ ] Right-click Start, **Terminal (Admin)**. Title must not say (x86). Over Quick Assist the admin prompt appears on the **person's** screen
- [ ] Paste the three lines. Enter, `1`, `1`. "Package downloaded" (or found), "Left", restart
- [ ] Sign-in screen: wait up to **2 minutes** for **Other user** (the join finishes in the background)
- [ ] They sign in with the new email, set a PIN (approve the "User OOBE" prompt if it appears)
- [ ] The "files are in place" message appears. They sign in to OneDrive, Outlook and Teams once

### After

- [ ] `dsregcmd /status | findstr /i "AzureAdJoined TenantName"`: YES, and the destination
- [ ] Add the person (or device) to your post-move Intune fixes: one script that sets the safety-net password and removes the old tenant's leftover admin, then Windows LAPS
- [ ] Phone: remove the old work account from Outlook, Teams and Authenticator, add the new one

### If something goes wrong

- **No "Other user" after 5 minutes:** sign in as `.\migadmin`, Terminal (Admin), `powershell -ExecutionPolicy Bypass -File C:/MigrationBackup/Move-TenantDevices.ps1 -Action MoveDevice` (retries the join)
- **No "files in place" message after 15 minutes:** `Get-Content C:/MigrationBackup/AutoCopy.log -Tail 20`, or run `-Action CopyProfile`
- **By hand:** Settings, Accounts, Access work or school, Connect, "Join this device to Microsoft Entra ID" with the person's new account. Works every time, but makes them a local admin

## Timings

The first laptop took most of a day: that's what the pilot is for. With option 9 the rest took **about 10 minutes each** of hands-on time, plus the person's first sign-in. One person can do a small office in a morning.

## Three ways to do it

| Route | How |
| --- | --- |
| Paid tool | Profile migration tools such as ForensiT User Profile Wizard re-point the existing Windows profile to the new account, so the desktop, settings and files stay where they are and the person sees no new profile. Paid editions can join the device to the destination Entra ID and run unattended. Some migration suites have a device agent that does the same. Test one laptop, end to end, before you buy |
| Microsoft's own | Windows Autopilot: export the device hashes, delete the devices from the source Autopilot, import them into the destination, then reset each device. Clean, and Intune builds it the way the destination wants, but the device is wiped, so the profile backup in the checklist matters even more. Free with Intune |
| This toolkit | `Move-TenantDevices.ps1` option 9: free, no wipe, one remote session per laptop, files copied at first sign-in |
| By hand | Leave in Settings, join in Settings with the person's account, copy their files as an admin. Free, slow, and the person ends up a local admin |

## Tips and traps

- **The current ADK designer crashes from the command line.** ICD.exe in ADK 10.1.26100 and later is missing `Microsoft.Diagnostics.Tracing.EventSource.dll` (FileNotFoundException). Option 1 runs a private copy with that file added from its NuGet package (1.1.28).
- **Windows PowerShell 5.1 + `2>&1` + `$ErrorActionPreference = 'Stop'` = a crash on any native stderr.** icacls always prints "Access is denied" for the legacy junctions in a profile (`Application Data`, `Cookies`). Relax the preference around native calls and verify the result instead.
- **"Profile still loaded" after sign-out is normal.** Switch user, a lingering session or a background process keeps the hive open. A restart, then signing in as the local admin only, clears it every time.
- **Do it in one sitting.** The staged route (option 4) needs a local-admin sign-in after every restart, and remote tools drop at each one. Option 9 does all the hands-on work in the person's own session, before the first restart, then a scheduled task (SYSTEM, at logon) copies their files when they first sign in with the new email. Technician time per laptop: about 10 minutes.
- **Give the script a link, not a file.** Upload the .ppkg once to OneDrive, share it as "Anyone with the link" with an expiry date (or an Azure blob SAS URL), and option 9 downloads it itself. Never publish the package: it joins any device to the tenant. Delete the link when the last laptop is done.
- **Send the package through Teams, not email.** Exchange's attachment filter blocks scripts and some archives; Teams files go through OneDrive. Option 9 finds the .ppkg in the person's Downloads.
- **Copy the person's OneDrive before you move their laptop.** With Files On-Demand most OneDrive files are cloud-only placeholders: they look present (full size in Explorer) but no data is on the disk, so nothing on the laptop can copy them once the old account is blocked. Only the tenant-to-tenant OneDrive copy brings them. Do it first, and the files are waiting at first sign-in.
- **Profile handover can end in a black screen.** In the field it mapped the profile correctly but the first sign-in hung on the per-user setup screens (UserOOBEBroker, Windows Hello) with a vendor agent crashing alongside. Copying into the fresh profile worked first time. Copy is now the default.
- **Look for duplicate destination accounts before copying.** An unlicensed `@tenant.onmicrosoft.com` twin of each person (from a set-up step run twice) makes tools guess the wrong OneDrive. Group users by display name, compare licences and addresses, prove which one the laptop uses (an Entra SID `S-1-12-1-a-b-c-d` is the account's object ID), then remove the twin.
- **Same email, different tenant: clear the cached sign-ins.** The handed-over profile still holds the old tenant's tokens (Microsoft.AAD.BrokerPlugin, Office `Common\Identity`, Teams cache). Left alone, Office and Teams loop on "Something went wrong". Option 5 renames or exports them. With option 9 the person gets a fresh profile, so this only matters for the handover route.
- **Entra profiles don't always show as `AzureAD\name`.** On current Windows they can show a tenant prefix (`CONTOSO\name`). Match them by SID instead: Entra user SIDs start `S-1-12-1-`.
- **Over Quick Assist, typing is where mistakes happen.** Every choice is a numbered list; irreversible steps need an explicit 1 (Enter alone does nothing). The clipboard doesn't cross either: send commands to the person in Teams and paste them in their session.
- **32-bit PowerShell can't see dsregcmd.** Some "Run as administrator" shortcuts open Windows PowerShell (x86). There `dsregcmd` is not found and BitLocker and registry calls are redirected. The script now relaunches itself in 64-bit; by hand, run `C:/Windows/Sysnative/WindowsPowerShell/v1.0/powershell.exe` first.
- **The Windows sign-in broker can't get this token.** The first-party apps allowed to request a device registration token have no broker redirect URI, so WAM fails with AADSTS50011. Option 1 uses a sign-in window instead.
- **Each bulk token request creates a `package_<guid>` account** in the destination, exactly as the designer does. Failed or repeated runs leave extra ones; delete the unused ones, never the one behind the package you are using (deleting it kills that package).
- **Device code is often blocked.** Many tenants block device code flow with Conditional Access. Option 1 signs in with a Microsoft sign-in window or the browser instead, so MFA and Conditional Access apply as normal. Use `-UseDeviceCode` only if you know it is allowed.
- **The account that joined the laptop is its admin, not your global admin.** On the laptops we moved, the source's global admin role had never been added to the local Administrators group; only the person who originally joined the laptop (and the built-in, disabled Administrator) had rights. Run `net localgroup administrators` first, and plan to use that account for the admin prompt.
- **Quick Assist hides the admin prompt from you.** UAC runs on the secure desktop, so the prompt appears only on the person's screen. Either they type the admin password, or you use a remote tool that handles UAC. `Start-Process -Credential` with an Entra account fails ("user name or password is incorrect"), so it is not a way round it. After the move, manage local admin through Intune (a platform script, then Windows LAPS), never by sharing passwords.
- **Scripts may be blocked.** Some laptops have PowerShell scripts disabled. Start the script with `powershell -ExecutionPolicy Bypass -File ...`; it affects that run only.
- **Wait for "Other user".** After the restart the join can still be finishing. Give the sign-in screen up to two minutes before reaching for the safety-net admin.
- **Type `.\migadmin`, not `migadmin`.** At an admin prompt or the sign-in screen on an Entra joined device, a bare name is treated as a work account.
- **One safety-net password for every laptop.** A random one per laptop gets lost. Pass `-LocalAdminPassword` (letters, digits and `!` only: a UK keyboard over a remote session swaps `@` and `"`), then let LAPS take the account over.
- **People with no Entra joined laptop may still have a computer.** One person had their own PC with a local Windows account. Ask everyone what they work on; option 9 moves stand-alone PCs too.
- **A device owner who isn't migrating.** The device list showed a laptop registered to someone who had left. Save their OneDrive and mail as former staff first, then reset the laptop (Remove everything) and let the next person join it at the setup screens with their work account.
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
| 2.0.0 | 2026-11-13 | Rewritten around option 9 (one go) after the live run: TL;DR, option table, per-device checklist and timings. Staged route and handover kept as fallbacks |
| 1.6.2 | 2026-11-11 | Move-TenantDevices 1.4.5 and 1.4.6: `-LocalAdminPassword` (one safety-net password on every laptop); option 9 also moves a stand-alone PC with a local Windows account (no leave, joins, copies the local profile at first sign-in). |
| 1.6.1 | 2026-11-11 | Move-TenantDevices 1.4.1 and 1.4.2: option 9's safety-net admin always gets a fresh easy password and is tested before leaving; option 9 downloads the package itself from a private expiring link (`-PackageUrl`, or paste it when asked). |
| 1.6.0 | 2026-11-11 | Move-TenantDevices 1.4.0: option 9 `-Action MoveInOneGo`, the recommended laptop route. One run while the person is signed in (finds the .ppkg in their Downloads, BitLocker key, safety-net admin, leave, join, restart); a SYSTEM task copies their files at first sign-in and removes itself. |
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
