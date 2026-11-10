# 10 Decommission the source (planned)

**TL;DR:** Planned. After cutover the source tenant is still there, still billing, still holding data and still accepting sign-ins unless you stop it. This part will cover closing it down on purpose: block sign-in, run the last incrementals, export and keep what must be kept, decide what happens to unlicensed OneDrives, remove licences, and close the tenant, all on a timeline set by the licence end date. Until the scripts land, this page is the plan.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Status | Planned. Block-SourceSignIn already exists in part 07 |
| Licence | MIT |

Part 10 of my [M365 tenant-to-tenant migration toolkit](../README.md).

## Why you need it

The source doesn't switch itself off. Left alone, people keep signing in to it (and saving files there), licences keep renewing, and when the subscription does lapse the data goes with it on Microsoft's schedule, not yours. Decommissioning is the step where you decide what's kept, where, and for how long, and get the business owner to sign that off in writing.

## What happens without it

- **People keep using the old tenant.** A bookmark, a phone app or a cached Teams login, and new work lands in a tenant you're about to close.
- **Data is lost by default.** When licences end, mailboxes and OneDrives enter retention and are then deleted. Anything not copied or exported is gone.
- **Legal holds are broken.** Data under hold or retention in the source can't simply be deleted. Closing the tenant without dealing with it is a compliance problem.
- **You pay twice.** Source licences renew for another year because nobody cancelled them.

## Intended scope

| Step | What | Notes |
| --- | --- | --- |
| 1 | Block source sign-in | Use [07 Block-SourceSignIn.ps1](../07-DomainCutover/) on T+1 for everyone who moved. It leaves accounts with admin roles alone: block those in the admin centre, last |
| 2 | Final incrementals | Mail, OneDrive, SharePoint and Teams, one last time, after sign-in is blocked so nothing new arrives |
| 3 | Export and retain | Anything the business must keep but didn't move: eDiscovery export to PST, or keep it under a hold until the retention period ends |
| 4 | Unlicensed OneDrives | Decide per person: copy into the archive library ([New-DriveCopyList](../copy-options/)), relicense briefly and copy, export with eDiscovery, or pay for Microsoft 365 Archive storage. Record the decision |
| 5 | Remove licences | Only after steps 2 to 4 are signed off. Cancel the subscription so it doesn't renew |
| 6 | Close the tenant | Remove remaining domains, delete users, then cancel subscriptions and close the account in the admin centre |
| 7 | Timeline | Work back from the source licence end date. Leave at least two weeks between "last incremental" and "licences end" |

## Three ways to do it

| Route | How |
| --- | --- |
| Paid tool | Most migration tools (ShareGate included) can run a last incremental and report what changed. Some archiving products can take a full export before the tenant closes |
| Microsoft's own | Microsoft Purview eDiscovery for exports and holds; Microsoft 365 Archive (pay as you go storage) for OneDrives and sites you must keep but don't need live; the admin centre for blocking sign-in, removing licences and closing the subscription |
| By hand | Block each user in the admin centre, download what's needed from each OneDrive, and cancel the subscription. Workable for a handful of people, with a checklist and a second pair of eyes |

## Gotchas already known

- **Trap: admin-role accounts can't be blocked by script.** Graph refuses to change accounts that hold privileged admin roles unless you hold a higher role. Block them in the admin centre, last, and keep one break-glass account until the very end.
- **Trap: blocking isn't instant.** Existing sessions keep working until their tokens expire. Revoke sessions as well if it matters.
- **Trap: unlicensed OneDrives go on a timer.** Depending on the tenant's unlicensed OneDrive policy they're made read-only, archived, or deleted. Archived ones can't be read by any copy tool until reactivated, which can take time and cost money.
- **Trap: deleting users deletes their OneDrive** after the retention period. Copy or export first.
- **Tip:** after an acquisition, leaving the old company's transport rules, connectors and odd policies behind is usually right. Rebuild only what the business still needs ([05](../05-MailConfig/)).
- **Tip:** get the business owner to sign a one-page decommission plan: what's kept, where, for how long, and the date the tenant closes. The [handover template](../docs/Handover-Template.md) has a section for it.

## When it lands

Scripts planned here: a decommission readiness report (who can still sign in, licences, holds, unlicensed OneDrives and their state), and a licence removal step that refuses to run while holds or uncopied OneDrives remain.
