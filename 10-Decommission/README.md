# 10 Decommission the source

**TL;DR:** After cutover the source tenant is still there, still billing, still holding data and still accepting sign-ins unless you stop it. `Get-DecommissionReadiness.ps1` reads the source and tells you, PASS or WARN, who can still sign in, what's still licensed and when it renews, what's on hold, which OneDrives are unlicensed or archived, which domains and mailboxes are left, and whether mail is still flowing. The dated checklist below closes the tenant on purpose, on a timeline set by the licence end date, with every keep-or-delete decision signed off by the business owner. Examples use fabrikam.com as the source, moving into the contoso tenant.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.0.1 (Get-DecommissionReadiness 1.0.0) |
| Released | 10 November 2026 |
| Licence | MIT |

Part 10 of my [M365 tenant-to-tenant migration toolkit](../README.md). It starts the day after the [cutover runbook](../docs/Cutover-Runbook.md) ends.

## Why you need it

The source doesn't switch itself off. Left alone, people keep signing in to it (and saving files there), licences keep renewing, and when the subscription does lapse the data goes with it on Microsoft's schedule, not yours. Decommissioning is the step where you decide what's kept, where, and for how long, and get the business owner to sign that off in writing.

## What happens without it

- **People keep using the old tenant.** A bookmark, a phone app or a cached Teams login, and new work lands in a tenant you're about to close.
- **Data is lost by default.** When licences end, mailboxes and OneDrives enter retention and are then deleted. Anything not copied or exported is gone.
- **Legal holds are broken.** Data under hold or retention in the source can't simply be deleted. Closing the tenant without dealing with it is a compliance problem.
- **You pay twice.** Source licences renew for another term because nobody turned off recurring billing.
- **A scanner or website form keeps mailing the old tenant,** and nobody reads it.

## Get-DecommissionReadiness.ps1

Read-only. Signs in to the source with Graph and Exchange Online (and SharePoint Online if you give `-SharePointAdminUrl`) and checks:

| Area | Check | WARN when |
| --- | --- | --- |
| Sign-in | Enabled member accounts | Anyone apart from you and `-Exclude` (break-glass) can still sign in |
| Licences | Licensed users, by SKU | Anyone still holds a licence |
| Licences | Each subscription: status, seats, next renewal or end date (Graph `directory/subscriptions`) | It renews or ends within `-RenewalWarningDays` (30), or is already in Warning, Suspended or LockedOut |
| Domains | Custom domains left | Anything other than onmicrosoft.com |
| Mailboxes | Count by type and total size | INFO only: what goes when the tenant goes |
| Holds | Litigation, In-Place and retention holds on mailboxes; organisation-wide holds; inactive mailboxes | Any |
| Mail flow | Message trace, last 7 days (`Get-MessageTraceV2`) | Anything delivered into, or sent from, the source |
| OneDrive | With `-SharePointAdminUrl`: OneDrives whose owner has no licence, their size, lock and archive state | Any with data, and any archived |

```powershell
.\Get-DecommissionReadiness.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com `
    -SharePointAdminUrl https://fabrikam-admin.sharepoint.com -Exclude breakglass@fabrikam.onmicrosoft.com
```

Needs `Microsoft.Graph.Authentication` and `ExchangeOnlineManagement` 3.x; `Microsoft.Online.SharePoint.PowerShell` for the OneDrive checks. Global Reader covers most of it; message trace needs Exchange Administrator (or a role with message tracking), and the OneDrive check SharePoint Administrator.

| File | Contents |
| --- | --- |
| `DecommissionReadiness-<date>.csv` | One row per check: Area, Check, Result (PASS, WARN, INFO, SKIP), Detail, Fix. [tools/Get-MigrationReport.ps1](../tools/) picks it up |
| `DecommissionDetail-Users-<date>.csv` | Every member account, enabled or not, excluded or not, with licences |
| `DecommissionDetail-Subscriptions-<date>.csv` | Every subscription with status, seats, consumed, created and next lifecycle date |
| `DecommissionDetail-Mailboxes-<date>.csv` | Every mailbox left, with size, holds and last logon |
| `DecommissionDetail-OneDrives-<date>.csv` | Every OneDrive, owner, licensed or not, size, lock and archive state |

## The checklist

Fill in the dates on day one. Work back from the **licence end date** (the script's `NextLifecycle`, or Billing, Your products in the admin centre). If the end date is too close for this timeline, buy a short extension: it's cheaper than losing data.

| When | Date | Who | Step | Done |
| --- | --- | --- | --- | --- |
| T+1 | | Lead | Block sign-in for everyone who moved: [07 Block-SourceSignIn.ps1](../07-DomainCutover/) `-WhatIf`, then for real. Admin-role accounts in the admin centre, last. Keep one break-glass Global Administrator. Revoke sessions if it must be immediate | [ ] |
| T+1 | | Lead | Run `Get-DecommissionReadiness.ps1`. Keep the CSV as the starting point | [ ] |
| T+7 | | Lead | Final incremental copies: mail, OneDrive, SharePoint, Teams. Sign-in is blocked, so nothing new arrives after this | [ ] |
| T+7 | | Lead and business owner | Re-run the readiness check. Go through every WARN together. Write the decision for each one in the decommission plan ([handover template](../docs/Handover-Template.md)) | [ ] |
| T+14 | | Lead | Export anything to keep that didn't move: Purview eDiscovery, a search per person or mailbox, export to PST. Inactive mailboxes and anything on hold first. Store the exports where the business owner agreed, and record where | [ ] |
| T+14 | | Lead | Mail flow WARNs: find whatever still sends to or from the source (scanners, website forms, apps with an old onmicrosoft.com address) and repoint it | [ ] |
| T+30 | | Lead | Holds released by their owner, or exported. Then remove licences from users. Unlicensed OneDrives start their timer now | [ ] |
| T+30 | | Device support | Every Entra joined laptop moved with [08 Move-TenantDevices.ps1](../08-Devices/) (state Done on each). Then **delete** the old device objects in the source: Entra ID devices, Intune managed devices and Autopilot registrations (option 3 > move removes those). Delete only; never wipe or retire | [ ] |
| T+30 | | Lead | Turn off recurring billing on every subscription (admin centre, Billing, Your products, the subscription, Recurring billing). Or cancel, if the agreement allows | [ ] |
| Licence end minus 14 days | | Lead and business owner | **Decide on archived and unlicensed OneDrives**, one by one: already copied (nothing to do), relicense the owner and copy (reactivates an archived one free in 24 to 48 hours), eDiscovery export, or pay-as-you-go reactivation. Record each decision | [ ] |
| Licence end minus 7 days | | Lead | Final readiness run. Only the WARNs the business owner accepted in writing should remain | [ ] |
| Licence end | | | Subscriptions expire, or you cancel them. Remove remaining custom domains first if any are left, or they stay tied to a tenant you can't easily get back into | [ ] |
| Licence end + 30 days (approx.) | | | **Expired**: admins can still sign in and get data out. Users can't | [ ] |
| Licence end + 120 days (approx.) | | | **Disabled**: admins lose access to the data. Last chance to reactivate the subscription | [ ] |
| Licence end + 180 days (approx.) | | | **Deprovisioned**: data deleted. The tenant itself can then be closed. Microsoft can take longer to remove the directory | [ ] |

The expired, disabled and deprovisioned periods depend on the agreement (direct, CSP, Enterprise Agreement, trial). Check the subscription's own dates in the admin centre; don't plan on the approximate figures above.

## Unlicensed OneDrives, in short

With Microsoft's default settings, a OneDrive whose owner has no licence becomes **read-only at day 60** and is **archived at day 93**. An archived OneDrive can't be read by any copy tool until it's reactivated. Two ways back: give the owner a licence again (free, 24 to 48 hours), or reactivate it with pay-as-you-go billing (a one-off charge per GB, then storage per GB per month). Either way, that takes time you won't have on the last day. Decide 14 days before the licence end, not after.

## Three ways to do it

| Route | How |
| --- | --- |
| Paid tool | Most migration tools (ShareGate included) can run a last incremental and report what changed. Archiving and backup products can take a full export of mailboxes and OneDrives before the tenant closes, and keep it searchable for years. Worth it when there's a long retention requirement |
| Microsoft's own | Microsoft Purview eDiscovery for holds and exports to PST; Microsoft 365 Archive and pay-as-you-go for OneDrives you must keep but don't need live; the admin centre for blocking sign-in, removing licences and turning off recurring billing. Free apart from pay-as-you-go storage |
| By hand | Block each user in the admin centre, download what's needed from each OneDrive, and cancel the subscription. Workable for a handful of people, with this checklist and a second pair of eyes |

## Tips and traps

- **Trap: admin-role accounts can't be blocked by script.** Graph refuses to change accounts that hold privileged admin roles unless you hold a higher role. Block them in the admin centre, last, and keep one break-glass account until the very end.
- **Trap: blocking isn't instant.** Existing sessions keep working until their tokens expire. Revoke sessions as well if it matters.
- **Trap: holds outlive licences.** Removing a licence or deleting a user doesn't release a litigation or retention hold; the data becomes an inactive mailbox. Deleting the tenant does release it. Get the hold owner's decision in writing.
- **Trap: wipe and retire reach real laptops.** A laptop that has left the source tenant can still have a pending Intune action waiting for it if it was ever offline when the action was sent, and a wipe or retire sent to the wrong record hits a laptop someone is using in the destination. For old device objects, delete the record only, never wipe or retire. Check `Get-DeviceInventory.ps1` against the destination first, so you know every laptop has arrived.
- **Trap: deleting users deletes their OneDrive** after the retention period. Copy or export first.
- **Trap: a renewal date you didn't see.** Annual subscriptions renew automatically. A CSP subscription is cancelled through the partner, not the admin centre. Find out which kind you have at T+1.
- **Trap: message trace reads 5,000 rows at most.** If the script says so, look at the rest in the Exchange admin centre before you call the source quiet.
- **Tip:** run `Get-DecommissionReadiness.ps1` at T+1, T+7, T+30 and 14 days before licence end. The WARN count should only go down.
- **Tip:** after an acquisition, leaving the old company's transport rules, connectors and odd policies behind is usually right. Rebuild only what the business still needs ([05](../05-MailConfig/)).
- **Tip:** get the business owner to sign a one-page decommission plan: what's kept, where, for how long, and the date the tenant closes. The [handover template](../docs/Handover-Template.md) has a section for it.

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 1.0.1 | 2026-11-10 | Checklist: delete old device objects in the source once every laptop has moved with 08; never wipe or retire |
| 1.0.0 | 2026-11-10 | First release of `Get-DecommissionReadiness.ps1`: sign-in, licences and renewal dates, domains, holds, mailboxes, mail flow and unlicensed OneDrives, PASS and WARN. Dated checklist from T+1 to tenant deletion. Replaces the planning page |
