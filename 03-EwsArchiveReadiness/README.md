# Test-EwsArchiveReadiness

**TL;DR:** Main mailboxes copying fine but every archive failing with 'Unauthorized'? It's almost always EWS, in one tenant or the other. This script signs in to the source and then the destination, checks every EWS layer that can block an archive copy, and fixes them with `-Fix`. One run, both tenants. It's read-only without `-Fix`. Examples use fabrikam.com as the source and contoso.com as the destination.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 2.0.2 |
| Released | 9 November 2026 |
| Licence | MIT |

Part 3 of my [M365 tenant-to-tenant migration toolkit](../README.md). Run it before you copy archives, once [New-TargetMailboxes](../02-TargetMailboxes/) has built the destination.

## Why you need this

Most migration tools now copy the main mailbox through Microsoft Graph. Many of them still read and write In-Place Archives over EWS (Exchange Web Services).

EWS changed in October 2026. Once a tenant has EWS on, only apps whose ID is on its `EwsAllowedAppIDs` list are let in. Microsoft pre-filled that list in many tenants from recent EWS activity, so expect to find IDs already there, and don't wipe them. Tenants that never set EWS are being switched off in phases. It's switched off for good on **1 April 2027**.

A tenant-to-tenant archive copy goes through EWS **twice**: once to read from the source, once to write to the destination. Both tenants have to let the app in, at every layer.

## What happens without it

It fails quietly, which is the worst kind of failure.

- **Main mail copies fine.** Graph isn't affected, so the first pass looks healthy and the reports look green.
- **Every mailbox with an archive fails,** usually within seconds, with 'Unauthorized access', 'Unable to access source archive mailbox' or something similar. Mailboxes without archives keep succeeding, which makes it look like a mailbox problem. It isn't.
- **Nothing in the error says which layer blocked it,** or which tenant. You end up checking permissions, licences and archive provisioning first, because that's what 'Unauthorized' suggests.
- **Cutover arrives with the archives still in the old tenant.** On a typical tenant, archives are most of the data. On the migration this came from, they were 1.25 TB of 1.7 TB.
- **After 1 April 2027 the EWS route is gone,** so leaving archives for later isn't an option.

I lost most of a Saturday to this on a live migration, working through the layers one at a time. The last thing I found was an old rule I'd set myself two days earlier, following advice written before October.

## What it checks, in each tenant

| # | Check | Why it blocks archives |
| --- | --- | --- |
| 1 | Tenant `EwsEnabled` | Off means no EWS for anyone |
| 2 | Tenant `EwsAllowedAppIDs` | Since October 2026 this is the list that counts. Your tool's app ID has to be on it, in **both** tenants |
| 3 | Old `EwsApplicationAccessPolicy`, `EwsAllowList`, `EwsBlockList` | These match **user agents**, not app IDs. Guides written before October tell you to set them. `EnforceAllowList` with an app ID in the list doesn't let your app in. It keeps it out |
| 4 | Per-mailbox EWS (`Set-CASMailbox`) | A mailbox with `EwsEnabled` False blocks EWS on its own, whatever the tenant says |
| 5 | `full_access_as_app` consent | The app needs the Exchange Online application permission to open other people's mailboxes over EWS |
| 6 | Archive readiness (`-CheckArchives`) | A new archive can show `ArchiveStatus` Active for hours before it can be read or written |

## Requirements

- `ExchangeOnlineManagement` 3.x. Use 3.7.2 or later if you can: the script uses `-DisableWAM` to give you a fresh account picker when Windows signs you in to the wrong tenant
- Exchange Administrator in each tenant for `-Fix`
- For the consent check: `Microsoft.Graph.Authentication` and `Microsoft.Graph.Applications`, with `Application.Read.All`. Skip it with `-SkipConsentCheck`
- Your migration tool's application (client) ID. Its documentation, or Entra ID > Enterprise applications, will have it
- Both tenant IDs. Entra admin centre > Overview, in each tenant

```powershell
Install-Module ExchangeOnlineManagement, Microsoft.Graph.Authentication, Microsoft.Graph.Applications -Scope CurrentUser
```

## Usage

```powershell
$p = @{
    AppId               = '<migration-tool-app-id>'
    SourceTenantId      = '<source-tenant-guid>'
    DestinationTenantId = '<destination-tenant-guid>'
}

# 1. Check both tenants. Changes nothing.
.\Test-EwsArchiveReadiness.ps1 @p

# 2. See what -Fix would change
.\Test-EwsArchiveReadiness.ps1 @p -Fix -WhatIf

# 3. Fix both tenants and check the archives are ready
.\Test-EwsArchiveReadiness.ps1 @p -Fix -CheckArchives

# Name the admin accounts, so Windows can't sign you in to the wrong tenant
.\Test-EwsArchiveReadiness.ps1 @p -SourceAdmin admin@fabrikam.onmicrosoft.com -DestinationAdmin admin@contoso.onmicrosoft.com

# One side only
.\Test-EwsArchiveReadiness.ps1 @p -Only Destination
```

You're asked to sign in twice: source first, then destination. If a sign-in lands in the wrong tenant, the script says so, disconnects and asks again before it touches anything.

After fixing, restart your migration tool so it gets fresh tokens, and run step 1 again to confirm. Microsoft says allow-list changes can take up to 24 hours. Mine worked within about half an hour, but plan for the longer figure.

## Output

Each finding prints under its tenant as OK, Fixed, Problem or Warning. A summary shows the problems, fixes and warnings per tenant. Everything goes to `EwsArchiveReadiness-<date>.csv` with a Side column. A `-Fix -WhatIf` run writes it too, so you have something to show before you change anything.

## Three ways to do it

| Route | How |
| --- | --- |
| Paid tool | Your copy tool's own pre-flight. ShareGate and others test connections and permissions, but on the live migration none of them said "EWS allow list" or which tenant. They said 'Unauthorized'. Use this to find the layer, then let the tool copy |
| Microsoft's own | Native cross-tenant mailbox migration doesn't use EWS at all: it's a mailbox move between the tenants, archives included. If EWS is the only thing stopping you and you have time to set up the [native route](../copy-options/A-Microsoft-native/), that's the way round it, and it keeps working after 1 April 2027 |
| By hand | `Get-OrganizationConfig \| Format-List Ews*` and `Get-CASMailbox -ResultSize Unlimited \| Where-Object EwsEnabled -eq $false` in each tenant, then `Set-OrganizationConfig -EwsAllowedAppIDs` with the full list (read it first). Check `full_access_as_app` in Entra ID, Enterprise applications, the app, Permissions. About 20 minutes per tenant once you know the layers |

## Tips and traps

- **Tip:** run it without `-Fix` first and keep the CSV. If anything else in the tenant stops working after the fix, it's your record of what was there before.
- **Tip:** test with one small archive before you queue the big ones. A pass in 10 minutes is better proof than any report.
- **Trap:** fixing one tenant. A copy reads from the source and writes to the destination. Both need the app ID on the list.
- **Trap:** `EwsAllowedAppIDs` with one ID replaces the whole list. Read it, add yours, write it back. The script does.
- **Trap:** guides written before October 2026. `EwsApplicationAccessPolicy EnforceAllowList` with an app ID matches user agents, not apps, and keeps your tool out.
- **Trap:** waiting on 1 April 2027. EWS goes for good. Anything still in a source archive after that needs the native route or another way out.

## Lessons learnt

- **Check both tenants, and check which one you're in.** It's easy to check the same tenant twice and think both are done. My source still had the old rule hours after I thought it was gone. That's why the script signs in to each tenant itself and checks the tenant ID.
- **Windows reuses your last sign-in.** If you work in both tenants on the same PC, `Connect-ExchangeOnline` can silently sign you in to the wrong one. Name the account with `-UserPrincipalName`, or use `-DisableWAM` to get the account picker.
- **`EwsAllowedAppIDs` replaces the whole list.** Read it, add your ID, write it back. The script does this. A one-liner with a single ID wipes everything else that was allowed.
- **Only one app on the list means only one app gets EWS.** If anything else in the tenant still uses EWS, such as a scanner, a CRM or an archiving add-on, add its app ID too. Check Admin centre > Reports > Usage > Exchange for EWS usage first.
- **The old user-agent settings are a trap.** Plenty of posts still show `-EwsApplicationAccessPolicy EnforceAllowList -EwsAllowList <app id>`. That's what I set first. It isn't what Exchange checks for apps now, and it actively blocks them.
- **Clear the old rules by clearing the policy.** `Set-OrganizationConfig -EwsApplicationAccessPolicy $null` removes the policy and both lists. Clearing a list on its own fails ('EwsApplicationAccessPolicy must be specified'), and so does clearing both lists in one call.
- **'Unauthorized' isn't always permissions.** My app had `full_access_as_app` the whole time. Check every layer before you re-consent anything.
- **Don't keep kicking new archives.** `Start-ManagedFolderAssistant` on brand-new archives mostly returned timeouts for me. Waiting worked better. `-CheckArchives` tells you when they're ready.
- **After 1 April 2027 none of this works.** If you have big archives to move, move them this winter.

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 2.0.2 | 2026-11-09 | The findings CSV is written under `-WhatIf` too, so a dry run leaves a report |
| 2.0.1 | 2026-11-09 | Missing `full_access_as_app` is a Warning, since tools using delegated EWS don't need it. Timing and wording match Microsoft's guidance (allow-list changes can take up to 24 hours) |
| 2.0.0 | 2026-11-08 | Checks source and destination in one run. Separate sign-in for each, with a wrong-tenant guard and a retry through a fresh account picker. `-SourceAdmin`, `-DestinationAdmin` and `-Only`. Summary per tenant |
| 1.0.4 | 2026-11-08 | Clears legacy rules by setting `EwsApplicationAccessPolicy` to `$null`, which removes both lists in one go |
| 1.0.3 | 2026-11-08 | An active legacy `EnforceAllowList` counts as a Problem |
| 1.0.2 | 2026-11-08 | Only reports Fixed when Exchange accepts the change. Reads list values whether they come back as a string or a collection |
| 1.0.1 | 2026-11-08 | `-TenantId` guard |
| 1.0.0 | 2026-11-08 | First release |
