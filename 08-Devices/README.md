# 08 Device rejoin (planned)

**TL;DR:** Planned. Laptops joined to the source tenant's Entra ID have to leave it and join the destination, and every app on them (Outlook, OneDrive, Teams) has to sign in again. This part will cover how to do that without losing anyone's files or a morning of their time: back up the profile, leave the source, join the destination, enrol in Intune, and sign the apps back in. Until the scripts land, this page is the plan and the traps I already know about.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Status | Planned. Scripts follow once they've been run on real devices |
| Licence | MIT |

Part 8 of my [M365 tenant-to-tenant migration toolkit](../README.md).

## Why you need it

A Windows device can be Entra-joined to one tenant only. After cutover the person's account in the source is blocked, so the device they log on to every morning still trusts a tenant that no longer knows them. Cached credentials keep them working for a day or two, then it falls apart: Windows Hello stops, OneDrive stops syncing, Outlook prompts forever, and Intune policy and apps from the source keep applying.

## What happens without it

- **People log on with cached credentials** and think all is well until the cache expires, usually on a Monday.
- **OneDrive keeps syncing to the old tenant**, or to nowhere, and files saved after cutover go to the wrong place.
- **The device stays managed by the source Intune**, so the source can still wipe it, and the destination can't manage it.
- **A local profile is lost** when the device is reset or rejoined without a backup, including desktop files that were never in OneDrive.

## Intended scope

| Step | What | Notes |
| --- | --- | --- |
| 1 | Inventory devices | From the source: join type (Entra joined, hybrid, registered), owner, OS, Intune, last sign-in. Part 01 records some of this |
| 2 | Back up the profile | Make sure Known Folder Move has synced Desktop, Documents and Pictures. Copy anything outside OneDrive (browser bookmarks, Outlook signatures, sticky notes, local PSTs) |
| 3 | Leave the source Entra ID | Settings, Accounts, Access work or school, Disconnect. Needs a local admin account on the device, created first |
| 4 | Join the destination | Settings, Accounts, Access work or school, Connect, "Join this device to Microsoft Entra ID", with the person's new sign-in |
| 5 | Intune enrolment | Automatic if MDM user scope is set in the destination. Check the device appears and policies apply |
| 6 | Re-sign in the apps | Outlook (new profile), OneDrive (sign in, pick the folder, let it re-sync), Teams, Office activation, browser profile |
| 7 | Phones | Remove the work account from Outlook, Teams and Authenticator, then add the new one. Intune app protection follows the new tenant |

## Three ways to do it

| Route | How |
| --- | --- |
| Paid tool | Commercial migration suites with a device agent can re-point Outlook profiles, OneDrive and the Windows profile remotely. Ask exactly what they do for Entra-joined devices (profile re-permissioning, not just Outlook) and test one laptop before you buy |
| Microsoft's own | Windows Autopilot: register the device hash in the destination, deregister it from the source, then reset. Clean, but the device is wiped, so the profile backup step matters even more. The free route for a fresh start |
| By hand | Disconnect and rejoin from Settings, as above, per device, with the person beside you. Twenty to forty minutes each. Fine for a small office, with a printed checklist |

## Gotchas already known

- **Trap: no local admin, no way back.** Disconnecting from Entra ID removes the only admin account many devices have. Create a local admin, and test it, before you disconnect.
- **Trap: a new profile.** Joining the destination creates a new Windows profile for the new account. The old one stays on disk but the person won't see it. Plan to copy what OneDrive didn't hold.
- **Trap: cached credentials hide the problem.** Blocking source sign-in doesn't stop a device logging on with cached credentials until the cache expires. Don't treat "they can still log on" as success.
- **Trap: BitLocker keys.** Recovery keys for Entra-joined devices are stored in the source tenant. Export them before you leave, or you lose the only copy when the source is closed.
- **Tip:** the OneDrive URL changes with the sign-in name. Sign OneDrive in fresh rather than trying to re-point the old sync relationship.
- **Tip:** do one device end to end before the night, and time it. That number sets how many people support can handle on day one.

## When it lands

Scripts planned here: device inventory from both tenants, BitLocker key export from the source, and a pre-flight check (local admin present, Known Folder Move synced, OneDrive healthy) to run on each device before it's rejoined.
