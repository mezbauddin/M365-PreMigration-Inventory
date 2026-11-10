# Risk register

**TL;DR:** The risks I see on most tenant-to-tenant migrations, with how likely they are, how much they hurt, what to do about them and which part of the toolkit helps. Start from this list, delete what doesn't apply, add what discovery turned up, and review it with the business owner before every go/no-go.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.2.0 |
| Released | 10 November 2026 |
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
| 25 | The DNS host isn't the registrar, and nobody can log in to it, so the TXT and MX can't be changed on the night | Medium | High | At T-14 check `Resolve-DnsName <domain> -Type NS`. If the host is out of reach, scan every record, recreate them at the registrar's DNS, verify, and switch name servers at least 48 hours before cutover | [07 Export-DomainDns -FullScan, Set-DomainDnsRecords](../07-DomainCutover/), [DNS move playbook](DNS-Move-Playbook.md) |
| 26 | Name server change still propagating on cutover night (NS records often cached for 48 hours), so half the internet sees old records | Medium | High | Move name servers days before, never on the day. Verify against the new name servers and a public resolver before cutover | [DNS move playbook](DNS-Move-Playbook.md) |
| 27 | Domain removal blocked by deleted users, proxy-only users, synced users or apps with URLs on the domain, found only on the night | High | High | Run `Remove-SourceDomain -BlockersOnly` at T-7 and again at T-1, then take the `-pre` snapshot and run `Test-CutoverReadiness` with it. Clear deleted users with `Clear-DeletedUserDomain`. Re-run next morning, because Exchange and Entra ID sync on their own schedule | [07](../07-DomainCutover/) |
| 28 | A script reuses an open session for the wrong account or tenant | Medium | High | Pass `-AdminUpn` to every tenant-changing script. It disconnects any other session and prints "Signed in as X to tenant Y" before changing anything | [02](../02-TargetMailboxes/), [04](../04-TargetOneDrives/), [05](../05-MailConfig/), [06](../06-GroupsTeamsSites/), [07](../07-DomainCutover/), [09](../09-Validation/) |
| 29 | The address snapshot is taken too early, so people and lists created since have no X500 or old aliases | Medium | Medium | Take a `-pre` snapshot at T-1 for the readiness check, and again on the night, immediately before the rollback point and `-RemoveDomain`. Use the night's file with `Switch-TargetAddresses`. `Test-CutoverReadiness` checks it | [07](../07-DomainCutover/) |
| 30 | Guests or other objects in the destination already hold source addresses, so old addresses can't be added to the real mailboxes | High | Medium | Find them before cutover (`Test-CutoverReadiness`), remove or rename the guests. 02b never maps to a guest | [02b](../02b-UserMapping/), [07](../07-DomainCutover/) |
| 31 | Location-based Conditional Access in the destination blocks the admin's test sign-in (error 53003), which looks like a failed migration | Medium | Medium | Plan who tests from where at T-1. Test from an allowed location, or exclude the test account for the night and remove the exclusion after. 09 lists every location policy | [09](../09-Validation/), [runbook](Cutover-Runbook.md) |
| 32 | Moving DNS copies an alias (flattened) apex record as a fixed IP address, and the website breaks when that IP changes | Medium | Medium | Ask the web host what the apex should point at. Use their documented record or the new provider's own flattening. Never publish an apex CNAME | [DNS move playbook](DNS-Move-Playbook.md), [07 Add-TargetDomain](../07-DomainCutover/) |
| 33 | Graph module version clash ("Could not load file or assembly") stops a script part way through | Medium | High | Run `Test-ModuleHealth` at T-14 and before the night. One version of every Microsoft.Graph module. A fresh window per task | [tools](../tools/) |
| 34 | A device leaves its tenant with no working local admin, so nobody can sign in to join it to the destination | Medium | High | `Move-TenantDevices` option 4 refuses to leave unless the safety-net local admin exists, is enabled and is an administrator. Sign in as it once before stage 2. Always pass `-SourceTenantId`, so it can't run on the wrong machine, such as your own admin PC | [08](../08-Devices/) |
| 35 | Profile handover fails, so the person signs in to an empty profile or a broken one | Low | High | Option 5 checks no profile is loaded, verifies the result and rolls back on any failure, leaving the old profile untouched. Fallback is option 6, which copies the data. Do one willing person end to end first | [08](../08-Devices/) |
| 36 | The bulk enrolment package (`.ppkg`) leaks and joins unknown devices to the destination for up to 180 days | Low | High | Keep it in one private place, never a shared drive or email. Shortest validity that covers the device window. Delete it when the last laptop is done, and remove the `package_<guid>` account it created in the destination | [08](../08-Devices/) |

## Project-specific risks

Add yours here, in the same format.

| # | Risk | Likelihood | Impact | Mitigation | Owner |
| --- | --- | --- | --- | --- | --- |
| | | | | | |

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 1.2.0 | 2026-11-10 | Risks 34 to 36 for devices: leaving a tenant without a local admin, profile handover failure, a leaked enrolment package |
| 1.1.1 | 2026-11-10 | Blocker and snapshot timing match the cutover runbook |
| 1.1.0 | 2026-11-10 | Risks 25 to 33 from a live cutover: DNS host access, name server TTL, removal blockers, wrong-tenant sessions, snapshot timing, guest address conflicts, Conditional Access on test sign-ins, alias records, module clashes |
| 1.0.0 | 2026-11-09 | First release |
