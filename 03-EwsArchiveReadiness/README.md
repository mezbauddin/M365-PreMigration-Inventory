# Test-EwsArchiveReadiness

**TL;DR:** Your main mailboxes copy fine but every archive fails with 'Unauthorized'? It's almost always EWS. Run this script in the source and the destination tenant with your migration tool's app ID. It checks the five things that block archive copies and fixes them with `-Fix`. It's read-only without `-Fix`.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.0.4 |
| Released | 8 November 2026 |
| Licence | MIT |

Part 3 of my [M365 tenant-to-tenant migration toolkit](../README.md). Run it before you copy archives, after [New-TargetMailboxes](../02-TargetMailboxes/) has built the destination.

## Why

Plenty of migration tools copy the main mailbox through Microsoft Graph and fall back to EWS for In-Place Archives. Since October 2026, EWS in Exchange Online is off by default, and only apps on the tenant's `EwsAllowedAppIDs` list get in. EWS goes away for good on 1 April 2027.

That leaves a nasty failure mode. Main mail copies, so everything looks healthy. Then every mailbox with an archive fails with 'Unauthorized' or 'Unable to access source archive mailbox'.

I lost most of a Saturday to it on a live migration, working through the layers one at a time. This script checks them all in one go.

## What it checks

| # | Check | Why it blocks archives |
| --- | --- | --- |
| 1 | Tenant `EwsEnabled` | Off means no EWS for anyone |
| 2 | Tenant `EwsAllowedAppIDs` | Since October 2026 this is the list that counts. Your tool's app ID has to be on it |
| 3 | Old `EwsApplicationAccessPolicy`, `EwsAllowList`, `EwsBlockList` | These match **user agents**, not app IDs. Guides written before October tell you to set them. They don't let your app in, and they can still keep it out |
| 4 | Per-mailbox EWS (`Set-CASMailbox`) | A mailbox with `EwsEnabled` False blocks EWS on its own, whatever the tenant says. Main mail still copies over Graph, so only the archive fails |
| 5 | `full_access_as_app` consent | The app needs the Exchange Online application permission to open other people's mailboxes over EWS |

With `-CheckArchives` it also confirms every active archive can actually be read. A new archive can show `ArchiveStatus` Active for hours before it's usable.

## Requirements

- `ExchangeOnlineManagement` 3.x. Exchange Administrator for `-Fix`
- For the consent check: `Microsoft.Graph.Authentication` and `Microsoft.Graph.Applications`, with `Application.Read.All`. Skip it with `-SkipConsentCheck`
- Your migration tool's application (client) ID. Its documentation, or Entra ID > Enterprise applications, will have it

```powershell
Install-Module ExchangeOnlineManagement, Microsoft.Graph.Authentication, Microsoft.Graph.Applications -Scope CurrentUser
```

## Usage

Run it once per tenant. Connect to the source first, then repeat for the destination.

```powershell
# Check only
.\Test-EwsArchiveReadiness.ps1 -AppId <your-app-id>

# See what -Fix would change
.\Test-EwsArchiveReadiness.ps1 -AppId <your-app-id> -Fix -WhatIf

# Fix it, and check the archives are ready
.\Test-EwsArchiveReadiness.ps1 -AppId <your-app-id> -Fix -CheckArchives

# Pin it to a tenant, so you can't fix the wrong side of the migration
.\Test-EwsArchiveReadiness.ps1 -AppId <your-app-id> -TenantId <tenant-guid> -Fix

# Only a few mailboxes
.\Test-EwsArchiveReadiness.ps1 -AppId <your-app-id> -Mailbox alex@contoso.com, docs@contoso.com
```

Give fixes 15 to 30 minutes to take effect, then run it again without `-Fix`, and restart your migration tool so it picks up fresh tokens.

## Output

Each finding prints as OK, Fixed, Problem, Warning or Info, and the lot goes to `EwsArchiveReadiness-<tenant>-<date>.csv`.

## Lessons learnt

- **Check both tenants, and check which one you're in.** The source has to let the tool read archives and the destination has to let it write them. I 'confirmed' the fix twice in the same tenant and spent hours on the wrong side. That's why `-TenantId` exists.
- **`EwsAllowedAppIDs` replaces the whole list.** Read it, add your ID, write it back. The script does this. A one-liner with a single ID wipes everything else that was allowed.
- **Only one app on the list means only one app gets EWS.** If anything else in the tenant still uses EWS, such as a scanner, a CRM or an archiving add-on, add its app ID too. Check Admin centre > Reports > Usage > Exchange for EWS usage first.
- **Clear the old rules by clearing the policy.** `Set-OrganizationConfig -EwsApplicationAccessPolicy $null` removes the policy and both lists. Clearing a list on its own fails ('EwsApplicationAccessPolicy must be specified'), and so does clearing both lists in one call.
- **The old user-agent settings are a trap.** Plenty of blog posts still show `-EwsApplicationAccessPolicy EnforceAllowList -EwsAllowList`. That's what I set first. It isn't what Exchange checks for apps now.
- **'Unauthorized' isn't always permissions.** My app had `full_access_as_app` the whole time. Check every layer, including per-mailbox EWS, before you re-consent anything.
- **Don't keep kicking new archives.** `Start-ManagedFolderAssistant` on brand-new archives mostly returned timeouts for me. Waiting worked better. `-CheckArchives` tells you when they're ready.
- **After 1 April 2027 none of this works.** If you have big archives to move, move them this winter.

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 1.0.4 | 2026-11-08 | Clears legacy rules by setting `EwsApplicationAccessPolicy` to `$null`, which removes both lists in one go |
| 1.0.3 | 2026-11-08 | An active legacy `EnforceAllowList` counts as a Problem. It blocks any app whose user agent isn't listed |
| 1.0.2 | 2026-11-08 | Legacy allow and block lists cleared in separate calls, since Exchange rejects both at once. Only reports Fixed when Exchange accepts the change |
| 1.0.1 | 2026-11-08 | `-TenantId`: Graph signs in to the named tenant, and an Exchange session for any other tenant is refused |
| 1.0.0 | 2026-11-08 | First release |
