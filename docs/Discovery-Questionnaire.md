# Discovery questionnaire

**TL;DR:** The questions I ask before I put a price or a date on a tenant-to-tenant migration. Most of the nasty surprises on a migration were sitting in one of these answers, waiting for someone to ask. Run [01 Source inventory](../01-SourceInventory/) alongside it: the script tells you what's there, and these questions tell you what the business actually needs.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.1.0 |
| Released | 10 November 2026 |
| Licence | MIT |

Part of the [docs](README.md) for my [M365 tenant-to-tenant migration toolkit](../README.md).

## Why it matters

The scripts tell you what's in the tenant. They can't tell you that the phones run through Teams, that a flow nobody remembers sends the invoices, or that the licence renews in three weeks.

**Without it:** you quote for mail and files, and spend the week after cutover rebuilding things nobody mentioned, unpaid.

## How I use it

I send it ahead of a call, then go through it together. Write the answers down and get them agreed. When something changes later, you've got a record of what you quoted against.

Where the answer is "I don't know", that's fine. Note who will find out, and by when. An unknown is a risk; put it in the [risk register](Risk-Register.md).

## 1. The tenants

- Which tenant is the source and which is the destination? Does the destination already have people working in it?
- Who are the Global Administrators in each? Can I have an admin account in both for the length of the project?
- Is there a break-glass account in each tenant, and who holds the details?
- Is either tenant managed by a reseller or another IT provider? Do they need to be involved or told?
- Is anyone else making changes to either tenant during the project?

## 2. Domains and DNS

- Which domains are moving? List every one, including old ones nobody uses but that still receive mail.
- Who controls DNS for each domain, and how quickly can they make a change? Do they work evenings or weekends?
- Where is DNS hosted? Can I get access, or will someone make changes for me on the night?
- Are there third-party records in SPF or other TXT records (marketing platforms, CRM, ticketing)?
- Is the moving domain the default domain in the source tenant?
- Is there a domain the destination already owns that can be used for staging addresses?

## 3. People and licences

- How many people, and how many of them are leaving or have already left?
- What licences does each tenant have, and when do they end or renew?
- Will licences overlap? Both tenants need to be paid for until cutover and for a while after.
- Which mailboxes belong to leavers? Keep as shared, archive somewhere, or don't move?
- Who's buying the destination licences, and when will they be in place?

## 4. Directory sync and identity

- Is either tenant synced from on-premises Active Directory (Entra Connect or Cloud Sync)?
- If the source is synced, can sync be turned off before cutover? Sign-in names and addresses of synced users can't be changed in the cloud.
- Do people use the same password across other systems tied to this tenant?
- Is single sign-on in place for other apps (accounting, CRM, HR)? Those will need repointing to the destination.

## 5. Devices

- How many laptops and desktops, and which operating systems?
- How are they joined: Entra joined, hybrid joined, domain joined, or not joined at all?
- Are they managed by Intune or another tool? Are there compliance policies people depend on?
- Phones and tablets: company owned or personal? Do they use Outlook, the built-in mail app, or both?
- Who will touch each device after cutover, and is that in the quote? Plan for this; it's often the biggest labour cost.

## 6. Security

- Is MFA in use? Which methods (Authenticator app, SMS, hardware keys)?
- Are there Conditional Access policies or security defaults in either tenant? Who will rebuild them in the destination?
- Will people need to register MFA again in the destination? Plan time for it on day one.
- Are there admin roles or privileged accounts that need recreating?

## 7. Retention, holds and compliance

- Are any mailboxes on litigation hold or in-place hold? (01 lists them.)
- Are there retention policies, retention labels or sensitivity labels in use?
- Are there legal or regulatory reasons to keep the source data for a set period?
- Does anyone need eDiscovery access to old data after the move?
- Who signs off that it's acceptable to switch the source off?

## 8. Mail flow and apps that send mail

- What sends mail as the company? Printers and scanners, line-of-business apps, website forms, CRM, accounting software.
- How do they send: SMTP AUTH with a username and password, a connector, a relay, or a third-party service?
- Who owns each one, and can they change its settings on cutover day?
- Are there mail flow rules (transport rules), connectors or third-party filtering in front of Exchange Online?

## 9. Shared mailboxes and permissions

- Which shared mailboxes are in use, and who uses each one?
- Who has Full Access, Send As or Send on Behalf to other people's mailboxes? (01 and [05](../05-MailConfig/) export these.)
- Are there rooms or equipment mailboxes, or Bookings pages?
- Any forwarding to outside addresses? Is it still wanted?

## 10. Teams and SharePoint

- How many Teams, and which ones actually matter? ([06](../06-GroupsTeamsSites/) exports them with sizes.)
- Private or shared channels in use?
- SharePoint sites outside Teams: intranet, document centres, old team sites?
- Anything heavily customised: Power Apps, Power Automate flows, custom lists, site designs? These don't move with the content.
- Are files shared with people outside the company? Those links will break.
- How much OneDrive data in total, and who are the big users?
- Is Teams chat history needed? Plan for it not moving; say so early.

## 11. Archives and sizes

- How big are the mailboxes and archives? (01's summary gives totals.)
- Are there auto-expanding archives, or archives over 100 GB?
- Will the copy tool you're planning to use copy archives over EWS? EWS in Exchange Online ends on 1 April 2027, so big archives need to move before then.
- Is there a PST problem: old mail sitting on laptops that people expect to come too?

## 12. Dates and deadlines

- Is there a hard deadline? Licence end date, contract end, company sale, office move?
- Are there dates to avoid: month end, payroll, busy season, key people on holiday?
- What's the preferred cutover window? Friday evening is common, but check someone is around on Monday morning to help.
- How long after cutover should the source stay available, read-only?

## 13. People and communication

- Who's the business owner I report to, and who makes the go/no-go call?
- Who owns communication to staff? Will they send the emails, or should I? ([Comms templates](Comms-Templates.md).)
- Who will people go to for help on day one, and is that person briefed?
- Are there VIPs, people who travel, or people working unusual hours who need a different plan?
- Are there external parties (clients, suppliers) who need telling about anything?

## 14. Things copy tools don't move

Each of these is explained, with three ways to handle it, in [What doesn't move](What-Doesnt-Move.md). For every "yes", write down who owns it: they're the person who'll recreate it.

**Teams chats**
- Do people rely on 1:1 or group chat history? How far back, and for what?
- Is "chats don't come across" acceptable, or is it a deal breaker that needs a paid tool and a budget?
- Is there a legal reason to keep chat history? If so, an eDiscovery export from the source before it closes.

**Planner**
- Which Teams or groups have Planner boards in use? Who owns each?
- Are any Premium plans (Planner with Project features) in use?

**Forms**
- Which forms are live and still collecting responses? Who owns them?
- Is any form linked from the website, an email signature or a printed QR code?

**Power Automate**
- How many flows are there, and who owns each? (List them from the admin side, not by asking.)
- Do any flows run as, or belong to, someone who has left?
- Which flows are business-critical (approvals, invoicing, onboarding)?

**Power Apps**
- Any canvas or model-driven apps? What data do they use: SharePoint lists, Excel, Dataverse?
- Do they use premium connectors that need licences in the destination?

**Power BI**
- Any Power BI workspaces, reports or dashboards in use? Who publishes them?
- Do the authors still have the original .pbix files? Are there gateways or scheduled refreshes?

**Loop**
- Are Loop workspaces or Loop pages in use? Which ones matter?

**Recordings**
- Are meeting recordings or their transcripts needed after the move? Do any have an expiry date set?

**OneNote**
- Which shared notebooks are in use, and where are they (OneDrive, a site, a Teams tab, only on a laptop)?

**Bookings**
- Is Microsoft Bookings in use? How many calendars, and is the booking link on the website?
- Are there future appointments booked through it?

**Phones**
- Is Teams Phone in use? Calling Plan, Operator Connect or Direct Routing, and who's the operator?
- How many numbers, auto attendants and call queues? Who records the greetings?
- What's the operator's lead time for moving numbers?

**Public folders**
- Are public folders in use? Any mail-enabled ones? Who uses them, and for what?

**Inactive mailboxes**
- Are there inactive mailboxes in the source (deleted leavers on hold)? Do they need keeping, and for how long?

**Enterprise apps and single sign-on**
- Which apps do people sign in to with their Microsoft account (SSO)? Who's the admin or vendor contact for each?
- Do any use automatic user provisioning?

**Sensitivity labels and encryption**
- Are sensitivity labels in use? Do any of them encrypt?
- Is encrypted mail (Message Encryption or "Do not forward") used with clients?
- How much encrypted content is there, and does it need to stay readable after the source closes?

**Rules, auto-replies and signatures**
- Do people rely on inbox rules, especially ones that forward or file mail?
- Are signatures managed centrally (a third-party tool or transport rule) or set by each person?
- Who'll be on holiday over cutover and need an auto-reply on the night?

**Outlook and devices**
- Will devices be rejoined ([08](../08-Devices/)) or just have Outlook, Office, Teams and OneDrive signed in again?
- How many people use Outlook desktop with a large mailbox that will need to re-cache?

**Shared calendars**
- Which shared calendars do teams depend on (rotas, holidays, room or vehicle calendars)?
- Any internet calendar subscriptions that matter?

**Mobile devices**
- Are phones enrolled in Intune or under app protection policies in the source?
- Who helps people re-add their work account and re-register MFA on day one?

## After the call

- Put every unknown and every "probably" into the [risk register](Risk-Register.md).
- Turn section 14's answers into the owner-recreates checklist in [What doesn't move](What-Doesnt-Move.md).
- Use the answers to choose a [copy option](../copy-options/).
- Draft the [cutover runbook](Cutover-Runbook.md) dates, and check them against section 12 and the licensing timeline in the runbook.

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 1.1.0 | 2026-11-10 | Section 14: discovery questions for everything in What doesn't move |
| 1.0.0 | 2026-11-09 | First release |
