# What doesn't move

**TL;DR:** Mail, files, SharePoint and Teams channel content move well with any of the three copy options. A long list of other things doesn't move at all, or moves badly: Teams chats, Planner, Forms, flows, apps, reports, Loop, recordings, Bookings, phone systems, public folders, inactive mailboxes, single sign-on, encrypted content, rules and signatures, Outlook profiles, calendar subscriptions and phones. This page lists each one, what goes wrong, who rebuilds it, and three ways to deal with it (paid tool, Microsoft native, by hand). Walk the business owner through it during discovery, not on the Monday after cutover.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.0.0 |
| Released | 10 November 2026 |
| Licence | MIT |

Part of the [docs](README.md) for my [M365 tenant-to-tenant migration toolkit](../README.md). The discovery questions for every item here are in section 14 of the [Discovery questionnaire](Discovery-Questionnaire.md).

## Why this page exists

The question that sinks the post-cutover week isn't "where's my email?" It's "where's my Planner board?", "why doesn't the expenses flow run?" or "the phones go to voicemail". None of those are in a mailbox or a document library, so no copy tool sees them. If nobody listed them at the start, they look like migration failures. Listed early, with an owner, they're just tasks.

## What happens without it

- **People find out on day one.** A Team arrives with its files and channels, and its Planner tab is empty. The trust you built over the migration goes in a morning.
- **Things fail quietly.** A flow owned by a source account stops when that account is blocked. Nobody notices until the monthly report doesn't arrive.
- **Data is lost for good.** Encrypted mail and files can only be opened by the tenant that encrypted them. Close the source and they're unreadable forever.
- **The quote was wrong.** Rebuilding a phone system or ten Power Apps is days of work that nobody priced.

## How to use it

1. Go through the summary table with the business owner during discovery. Mark each row In use / Not used / Don't know.
2. For every row In use, name an owner. Usually that's the person who built it, not IT.
3. Put each one in the [risk register](Risk-Register.md) and the [cutover runbook](Cutover-Runbook.md) (most belong in T+2 to T+6).
4. Send the owner-recreates checklist below to each owner at T-7, with the date it has to be done by.

## Summary

| Item | Moves with a copy tool? | Who usually recreates |
| --- | --- | --- |
| [Teams 1:1 and group chats](#teams-11-and-group-chats) | No natively. Some paid tools, with limits | Nobody. People save what matters |
| [Planner](#planner) | Some paid tools, with a Team | Plan owner |
| [Forms](#forms) | No | Form owner |
| [Power Automate flows](#power-automate-flows) | No | Flow owner |
| [Power Apps](#power-apps) | No | App owner |
| [Power BI](#power-bi) | No | Report owner |
| [Loop](#loop) | Partly, as files | Workspace owner |
| [Stream (on SharePoint) recordings](#stream-on-sharepoint-recordings) | Yes as files, without the extras | Nobody, unless transcripts matter |
| [OneNote](#onenote-quirks) | Yes, with quirks | Each person reopens |
| [Bookings](#bookings) | No | Bookings admin |
| [Teams Phone, auto attendants, call queues](#teams-phone-auto-attendants-and-call-queues) | No | IT, with the phone provider |
| [Public folders](#public-folders) | Some paid tools | IT |
| [Inactive mailboxes](#inactive-mailboxes) | No, not as they are | IT, with whoever owns retention |
| [Enterprise apps and SSO](#enterprise-apps-and-single-sign-on) | No | IT, with each app vendor |
| [Sensitivity labels and encrypted content](#sensitivity-labels-and-encrypted-content) | Labels no; encrypted content becomes unreadable | IT, before the copy |
| [Inbox rules, auto-replies, signatures](#inbox-rules-auto-replies-and-signatures) | Depends on the tool | Each person, with a checklist |
| [Outlook profiles and cached credentials](#outlook-profiles-and-cached-credentials) | n/a, on the device | Day-one support |
| [Shared calendar subscriptions](#shared-calendar-subscriptions) | No | Each person |
| [Mobile devices](#mobile-devices) | n/a, on the device | Each person, with support |

## The owner-recreates checklist

Send this to every owner at T-7. One line per item they own.

| # | Item | Owner | Export or note taken before T-1 | Recreated in destination | Tested | Old one switched off |
| --- | --- | --- | --- | --- | --- | --- |
| 1 | [e.g. Expenses approval flow] | [name] | [ ] | [ ] | [ ] | [ ] |
| 2 | [e.g. Sales pipeline Planner board] | [name] | [ ] | [ ] | [ ] | [ ] |
| 3 | [e.g. Main line auto attendant] | [name] | [ ] | [ ] | [ ] | [ ] |

**Tip:** "Old one switched off" matters. A flow left running in the source after cutover keeps sending emails from a tenant that's being closed, and nobody knows why.

## Teams 1:1 and group chats

**What doesn't move:** chats aren't in a Team, so no Team or SharePoint copy sees them. Microsoft's native cross-tenant migration doesn't move them. The copies of chat messages kept in each person's mailbox for compliance are hidden and can't be turned back into chats. Meeting chats go the same way.

**Limits of the paid tools that do it:** usually recent months only, often posted as HTML or as messages from a migration account rather than the original sender, and reactions, edits, inline images, Loop components and threading are lost or flattened. Price it per user before you promise it.

**Owner recreates:** nothing to rebuild. People save what matters before T-1: files shared in chat are already in their OneDrive (they copy with it), decisions and instructions go into a document or a Team channel.

- **Paid tool:** some migration suites copy 1:1 and group chats with the limits above. Ask for a test with a real busy chat before you buy.
- **Microsoft native:** an admin can export chats with Purview eDiscovery (to keep a record, not to restore as chats). The Microsoft Graph export APIs can read chat history for a custom archive, with metered costs.
- **By hand:** tell people at T-7 (the [comms templates](Comms-Templates.md) already do). Copy anything important into a channel post or a document.

**Trap:** don't let anyone promise "chats will come across" in the sales conversation. It's the most common disappointment after a tenant move.

## Planner

**What doesn't move:** plans, buckets, tasks, checklists, assignments and comments live in the Planner service, not in the group's site. Copy a Team's files and channels and the Planner tab arrives empty or missing. Premium plans (Planner with Project features) live in Dataverse and are a separate job.

**Owner recreates:** the plan, buckets and open tasks. Closed tasks rarely need moving; export them to Excel for the record.

- **Paid tool:** some tools copy basic plans along with a Team (tasks, buckets, assignments). Check what they do with comments, attachments and completed tasks.
- **Microsoft native:** Planner's own "Export plan to Excel" for a record. The Graph Planner API can read and recreate plans with a script.
- **By hand:** recreate the plan in the destination Team and copy open tasks across. Fine for a board with thirty tasks.

**Tip:** recreate plans in the destination Team at T+1, once people can sign in, so tasks are assigned to real destination accounts.

## Forms

**What doesn't move:** forms belong to a person or a group in Microsoft Forms, not to OneDrive. Responses go with them. Links people have bookmarked or embedded on a website stop working when the source closes.

**Owner recreates:** each live form, then updates every place its link is published. Exports the responses to Excel first.

- **Paid tool:** I haven't found one that moves Forms reliably between tenants. Check with your vendor.
- **Microsoft native:** "Open in Excel" for the responses. "Share as a template" gives a link the owner can open from the destination account to duplicate the form, where the tenants' settings allow it.
- **By hand:** rebuild the form in the destination. Ten questions take ten minutes.

**Trap:** forms linked from a public website keep collecting responses in the source until someone changes the link. Put the website owner on the checklist.

## Power Automate flows

**What doesn't move:** flows live in a Power Platform environment, owned by a person, and use connections signed in as that person. Nothing copies them. When the owner's source account is blocked at T+1, the flow stops.

**Owner recreates:** each flow, its connections, and anything it points at (SharePoint sites, lists and mailboxes now have new addresses).

- **Paid tool:** a few tools move flows as packages. The connections still need signing in again.
- **Microsoft native:** export each flow as a package (.zip) from the source, import it in the destination, then create fresh connections and repoint the actions. Solution-aware flows move better as a solution export. The Power Platform admin PowerShell lists every flow and its owner.
- **By hand:** rebuild it. Often quicker for a five-step flow than fixing an import.

**Trap:** flows owned by someone who has already left are the ones nobody remembers. List flows by owner from the admin side, not by asking people.
**Tip:** turn the source flow off once the destination one runs. Two copies of an approval flow means two approvals.

## Power Apps

**What doesn't move:** canvas apps, model-driven apps, their data connections and any Dataverse tables.

**Owner recreates:** the app, its connections and its data source. Apps built on SharePoint lists need repointing to the copied list in the destination.

- **Paid tool:** rarely covered. Ask.
- **Microsoft native:** export the canvas app as a package and import it in the destination, then fix connections. Model-driven apps and Dataverse data move as solutions and with data import tools.
- **By hand:** rebuild. For anything with Dataverse, budget it as its own project.

**Trap:** premium connectors and Dataverse need licences in the destination. Check before you quote.

## Power BI

**What doesn't move:** workspaces, reports, semantic models (datasets), dashboards, scheduled refreshes, gateways and sharing. None of it is in SharePoint.

**Owner recreates:** downloads each report (.pbix) where downloads are allowed, republishes it to a new workspace, re-enters data source credentials, re-creates refresh schedules and sharing.

- **Paid tool:** specialist tools exist. Few general migration tools do it.
- **Microsoft native:** Power BI Desktop to download and republish. The Power BI REST APIs can export and import reports at scale.
- **By hand:** republish from the original .pbix files, if the author still has them.

**Trap:** reports whose data source is a SharePoint list or Excel file in the source break silently when the source closes. Repoint them at the copied file.

## Loop

**What doesn't move:** Loop app workspaces and pages live in SharePoint Embedded containers that copy tools mostly can't see. Loop components in chats and emails are `.loop` files in the sender's OneDrive.

**Owner recreates:** important workspaces as pages in the destination, or as documents.

- **Paid tool:** check carefully. Support for Loop workspaces is new and patchy.
- **Microsoft native:** `.loop` files in OneDrive copy as files and open again, but links to them from old chats and emails point at the source. Copy page content into a new destination workspace or a Word document.
- **By hand:** copy and paste the pages that matter.

## Stream (on SharePoint) recordings

**What doesn't move well:** meeting recordings are video files in the organiser's OneDrive (Recordings folder) or in the channel's SharePoint folder. They copy as files. Transcripts, chapters, view counts, the link from the meeting chat and recording expiry settings may not come with them.

**Owner recreates:** nothing in most cases. Anyone who needs a transcript downloads it before T-1.

- **Paid tool:** copies the video files with the OneDrive or site. Check transcripts.
- **Microsoft native:** download the transcript from the recording's page in the source. The video copies with OneDrive or SharePoint.
- **By hand:** download the recordings that matter.

**Trap:** recordings with an expiry date set in the source may already be gone by the time you copy. List the ones that matter during discovery.

## OneNote quirks

**What moves:** notebooks in OneDrive and SharePoint are folders of files, and copy tools move them.

**Quirks:** a notebook copied while people still have it open in the desktop app can end up with conflicts. Links between pages and sections often still point at the source. Every person has to close the old notebook in each OneNote app and open the new one. Notebooks pinned as Teams tabs need re-adding. Notebooks stored only on a laptop don't move at all.

**Owner recreates:** each person reopens their notebooks from the new location. Teams tab owners re-add the tab.

- **Paid tool:** most copy notebooks as files. Ask what they do with internal links.
- **Microsoft native:** in OneNote, close the notebook, then open it from the destination OneDrive or site.
- **By hand:** for a local-only notebook, move it to OneDrive in the source before the copy, so it travels.

**Tip:** ask people to sync OneNote and close it at T-1, in the T-1 email. Fewer conflicts.

## Bookings

**What doesn't move:** a Bookings calendar is a special scheduling mailbox with its configuration: services, staff, business hours, booking page and customer list. Copy tools don't recreate it as a working Bookings page. The public booking URL changes.

**Owner recreates:** the Bookings calendar, services, staff, hours and questions in the destination, then updates the link on the website and email signatures. Future appointments are re-entered or customers are asked to rebook.

- **Paid tool:** I don't know of one that moves a working Bookings calendar.
- **Microsoft native:** Bookings web app in the destination. Export the source's appointment list where the admin centre offers it, for reference.
- **By hand:** screenshots of every setting page in the source, then rebuild. An hour per calendar.

## Teams Phone, auto attendants and call queues

**What doesn't move:** phone numbers, resource accounts, auto attendants, call queues, holiday schedules, greetings, emergency addresses and calling policies. Voicemail is in the mailbox, so it copies with mail.

**Owner recreates:** IT, with the phone provider. Numbers move by the route they came in: Microsoft Calling Plan numbers through a Microsoft support or porting request between tenants, Operator Connect numbers through the operator, Direct Routing numbers by repointing the SBC.

- **Paid tool:** none move the configuration. The phone operator does the number side.
- **Microsoft native:** export each auto attendant and call queue's settings with Teams PowerShell (`Get-CsAutoAttendant`, `Get-CsCallQueue`) as your record, then rebuild them in the destination Teams admin centre.
- **By hand:** screenshots of every page, recordings of every greeting, then rebuild.

**Trap:** number moves have lead times measured in days or weeks and their own cutover window, and they rarely line up with the domain move. Plan a separate phone cutover, agreed with the operator, and tell people the main number may be on a temporary divert for a few hours.

## Public folders

**What doesn't move:** Microsoft's native cross-tenant migration doesn't move public folders. Many copy tools don't either. Mail-enabled public folders also hold addresses on the moving domain, which become blockers at release.

**Owner recreates:** IT, usually as a shared mailbox or a Microsoft 365 group in the destination rather than public folders again.

- **Paid tool:** some migration suites copy public folders between tenants. Check permissions and mail-enabled folders.
- **Microsoft native:** within the source, migrate public folders to a Microsoft 365 group first, then copy the group. Or export to PST from Outlook and import.
- **By hand:** for a few folders, drag them into a shared mailbox in Outlook, in the source, before the copy.

**Trap:** mail-enabled public folders hold addresses on the moving domain, and the release has to strip them like any other recipient. List them with `Get-MailPublicFolder` during discovery and decide at T-7 whether they're being moved, replaced or dropped.

## Inactive mailboxes

**What doesn't move:** an inactive mailbox is a leaver's mailbox that was deleted while on hold or under a retention policy. It isn't a recipient, so copy tools and 02 don't see it. It disappears when the source tenant closes.

**Owner recreates:** IT, with whoever owns retention. Decide per mailbox: keep the source open, export it, or bring it into the destination.

- **Paid tool:** copy it once it's restored to a normal mailbox (below).
- **Microsoft native:** restore it into a new or existing mailbox in the source (`New-MailboxRestoreRequest`, or recover it as a new mailbox), then copy it like any other leaver. Or export it to PST with Purview eDiscovery.
- **By hand:** eDiscovery export to PST and keep the files somewhere controlled, with a note of why and for how long.

**Trap:** 01 lists mailboxes on hold, but not inactive ones. `Get-Mailbox -InactiveMailboxOnly` in the source during discovery.

## Enterprise apps and single sign-on

**What doesn't move:** every app that signs people in with the source Entra ID (accounting, HR, CRM, ticketing): enterprise applications, SAML certificates, OIDC app registrations, user provisioning (SCIM) and who's assigned to each.

**Owner recreates:** IT, with each app's vendor or admin. In the destination: create the enterprise app, new certificate or secret, assign users, set up provisioning. In the app: swap the identity provider metadata. Check people map to the same account in the app (usually by email, which stays the same).

- **Paid tool:** none do this for you.
- **Microsoft native:** Entra admin centre, Enterprise applications, in both tenants. Gallery apps have step-by-step guides. Export the source's app list and its sign-in logs to find what's really used.
- **By hand:** a spreadsheet per app: vendor contact, SSO type, who's assigned, cutover slot. Many vendors need a ticket a week ahead.

**Trap:** an app registration with an identifier URI on the moving domain is a blocker at release. `-FixApplications` moves the URI, but anything that calls the app by that URI breaks. Find those at T-7.
**Tip:** sign-in logs from the last 30 days show which apps are actually used. The list is usually shorter than people fear.

## Sensitivity labels and encrypted content

**What doesn't move:** labels are defined per tenant. A labelled file keeps the label's metadata, but the destination doesn't know that label. Worse, anything **encrypted** by a label, by Purview Message Encryption or by older Rights Management, is protected with the source tenant's keys. In the destination it may not open, and once the source is closed nobody can open it.

**Owner recreates:** IT, before the copy. Recreate labels and policies in the destination. Decide what happens to encrypted content: decrypt it in the source before copying, or keep the source open (and licensed) as long as it's needed.

- **Paid tool:** some tools warn about encrypted items or can't copy them. None can re-encrypt with the destination's keys for you.
- **Microsoft native:** in the source, an admin with the Rights Management super user feature can open and remove protection, for example with the Purview Information Protection client and its PowerShell (`Set-FileLabel` to remove labels and protection). Then copy. Recreate labels in the destination Purview portal.
- **By hand:** owners open each protected file and save it without protection. Fine for a handful.

**Trap:** encrypted emails sent to people outside the company are readable through a link served by the source. Close the source and those links stop working for the recipients.

## Inbox rules, auto-replies and signatures

**What moves depends on the tool.** Microsoft's native cross-tenant migration moves the mailbox whole, so server-side rules and settings come with it. EWS and PST routes often don't copy rules, auto-replies or roaming signatures. Rules that forward to an address or move to a folder by its old ID can break even when copied. Client-only rules and signatures stored on the PC never move.

**Owner recreates:** each person checks their rules, auto-reply and signature on day one, using a list IT exported.

- **Paid tool:** check rules, auto-replies and signatures in the test batch. Some tools do some of them.
- **Microsoft native:** export with `Get-InboxRule -Mailbox <name>` and `Get-MailboxAutoReplyConfiguration` in the source, recreate with `New-InboxRule` and `Set-MailboxAutoReplyConfiguration` in the destination.
- **By hand:** screenshots of the rules screen before T-1. Add "check your rules and signature" to the cutover-morning email.

**Tip:** people who are away over cutover need their auto-reply set in the destination on the night, or senders hear nothing.

## Outlook profiles and cached credentials

**What doesn't move:** this lives on each PC: the Outlook profile, the offline cache, saved passwords in Windows Credential Manager, the Office "work or school" account, Teams' cache and the OneDrive sync link. After cutover they all point at the old tenant.

**Owner recreates:** day-one support, person by person, with a script for each step.

- **Paid tool:** some migration suites have a desktop agent that resets Outlook and OneDrive for the user.
- **Microsoft native:** in Outlook, a new profile (Control Panel, Mail, Show Profiles, Add). In Office, File, Account, sign out of the old account. OneDrive, Settings, Unlink this PC. Teams, sign out and clear the cache.
- **By hand:** Credential Manager, remove entries for the old tenant. Then the steps above. Ten minutes per person, longer for anyone with a big mailbox to re-cache.

**Trap:** device rejoin in [08](../08-Devices/) clears most of this for you. If devices aren't being rejoined, budget the time per person.

## Shared calendar subscriptions

**What doesn't move:** calendars other people shared with you and you added to Outlook, and internet calendar subscriptions (holiday lists, booking systems, a supplier's calendar). Both point at the source. Calendar folder permissions inside a mailbox move with native migration and some tools, but the "added calendar" entries in each person's Outlook don't.

**Owner recreates:** each person re-adds the calendars they use. The calendar's owner reshares it if permissions didn't copy.

- **Paid tool:** may copy folder permissions. Doesn't re-add calendars to other people's Outlook.
- **Microsoft native:** `Get-MailboxFolderPermission <mailbox>:\Calendar` in the source to record who had what; `Add-MailboxFolderPermission` in the destination. People re-add with Outlook, Add calendar, From directory.
- **By hand:** a list of who looks at whose calendar, from the team. Then reshare.

## Mobile devices

**What doesn't move:** the work account on each phone and tablet: Outlook mobile, Teams, OneDrive, the built-in mail app's account, Authenticator registrations for MFA, and any Intune enrolment or app protection policy from the source.

**Owner recreates:** each person, with support: remove the work account, add it again, register MFA in the destination. Company devices enrolled in the source Intune need retiring there and enrolling in the destination.

- **Paid tool:** none move this. Intune migration tools exist for policies, not for the phone itself.
- **Microsoft native:** in the source, retire or selectively wipe devices once people are on the destination. In the destination, set up Intune enrolment and app protection before T-1.
- **By hand:** the steps in the cutover-morning email. Five minutes per phone, longer if MFA registration goes wrong.

**Trap:** a phone still signed in to the source keeps delivering mail from the source mailbox until sign-in is blocked at T+1. People think it's working, then it stops.

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 1.0.0 | 2026-11-10 | First release: nineteen things copy tools don't move, with owner-recreates checklist and three ways for each |
