# Risk register

**TL;DR:** The risks I see on most tenant-to-tenant migrations, with how likely they are, how much they hurt, what to do about them and which part of the toolkit helps. Start from this list, delete what doesn't apply, add what discovery turned up, and review it with the business owner before every go/no-go.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.0.0 |
| Released | 9 November 2026 |
| Licence | MIT |

Part of the [docs](README.md) for my [M365 tenant-to-tenant migration toolkit](../README.md).

## How to read it

Likelihood and impact are High, Medium or Low, from my experience on small and mid-sized moves. Yours will differ. Change them once you've done discovery, and add an owner column if more than one person is working on the project.

## Common risks

| # | Risk | Likelihood | Impact | Mitigation | Toolkit part |
| --- | --- | --- | --- | --- | --- |
| 1 | Source won't release the domain on the night because something still uses it (deleted user, guest, group alias) | High | High | Run `Remove-SourceDomain -WhatIf` days before. Clear every blocker it lists, including the recycle bin | [07](../07-DomainCutover/) |
| 2 | Mail bounces or queues between the domain leaving the source and MX pointing at the destination | Medium | High | Lower TTLs a week before. Run release, add and verify back to back. DNS owner on call | [07](../07-DomainCutover/), [runbook](Cutover-Runbook.md) |
| 3 | Replies to old mail bounce with IMCEAEX errors | High | Medium | Snapshot the source before release. Add X500 addresses at cutover. Validate them | [07](../07-DomainCutover/), [09](../09-Validation/) |
| 4 | Archives fail to copy because the tool uses EWS and the tenants block it | Medium | High | Check and fix EWS settings in both tenants before the first copy | [03](../03-EwsArchiveReadiness/) |
| 5 | EWS retirement on 1 April 2027 stops archive copies part way | Medium | High | Move big archives first, and well before the date. Consider a copy option that doesn't use EWS | [03](../03-EwsArchiveReadiness/), [copy options](../copy-options/) |
| 6 | Shared mailboxes arrive but nobody can open them | High | High | Export and rebuild Full Access, Send As and Send on Behalf before cutover | [05](../05-MailConfig/) |
| 7 | Distribution lists missing, or missing members | Medium | Medium | Rebuild lists from the export. Compare member counts | [05](../05-MailConfig/), [09](../09-Validation/) |
| 8 | Teams and groups created by the copy tool on the wrong domain or with the wrong site URL | Medium | Medium | Create groups first, on your chosen domain, then copy into them | [06](../06-GroupsTeamsSites/) |
| 9 | Teams content copies but members don't | Medium | High | Import the user mapping into the copy tool before starting. Add members up front | [06](../06-GroupsTeamsSites/) |
| 10 | Source is still synced from on-premises, so sign-in names and addresses can't be changed | Low | High | Find out in discovery. Plan to turn sync off well before cutover | [Discovery](Discovery-Questionnaire.md) |
| 11 | Someone can't sign in on day one (still staged, unlicensed, wrong sign-in name) | Medium | Medium | Validate the night before and on the night. Day-one support in place | [02](../02-TargetMailboxes/), [09](../09-Validation/) |
| 12 | Not enough licences in the destination on cutover night | Low | High | Count seats against the plan at T-7 and at go/no-go | [02](../02-TargetMailboxes/) |
| 13 | OneDrive copy fails because OneDrives don't exist yet | High | Medium | Request OneDrives as soon as people are licensed | [04](../04-TargetOneDrives/) |
| 14 | Scanners, apps or website forms stop sending mail | High | Medium | List everything that sends mail in discovery. Owners ready to change settings on day one | [Discovery](Discovery-Questionnaire.md) |
| 15 | A wrong mail flow rule blocks or redirects mail in the destination | Low | High | Rebuild transport rules by hand, disabled, test, then enable | [05](../05-MailConfig/) |
| 16 | Data under legal hold or retention is lost when the source is switched off | Low | High | List holds in the inventory. Get written sign-off before decommissioning | [01](../01-SourceInventory/), [handover](Handover-Template.md) |
| 17 | Admin runs a command in the wrong tenant | Medium | High | Use the scripts' `-TenantId` guard. Disconnect between tenants. Check `Get-ConnectionInformation` before anything destructive | All |
| 18 | Laptops and phones take longer to reconnect than planned | High | Medium | Agree in the quote who touches each device. Print sign-in instructions | [Comms](Comms-Templates.md) |
| 19 | People lose Teams chat history and weren't told | High | Low | Say so in the T-7 email, plainly | [Comms](Comms-Templates.md) |
| 20 | Shared file links break | High | Low | Warn people in the FAQ. Reshare important links after the move | [Comms](Comms-Templates.md) |
| 21 | Temporary passwords leak | Low | High | Hand out in person or by phone. Delete the passwords file after use | [07](../07-DomainCutover/) |
| 22 | Source licences end before the copy finishes | Medium | High | Check end dates in discovery. Keep source licences for a few weeks after cutover | [Discovery](Discovery-Questionnaire.md) |
| 23 | MFA and Conditional Access not rebuilt, so the destination is less secure than the source | Medium | High | Rebuild security settings deliberately before cutover. Don't copy them blind | [Discovery](Discovery-Questionnaire.md) |
| 24 | Inventory CSVs or reports end up in email or a shared drive | Medium | Medium | Keep outputs on the admin machine. The repo's `.gitignore` excludes `*.csv` | All |

## Project-specific risks

Add yours here, in the same format.

| # | Risk | Likelihood | Impact | Mitigation | Owner |
| --- | --- | --- | --- | --- | --- |
| | | | | | |
