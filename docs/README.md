# Docs

**TL;DR:** The paperwork that goes around the scripts. A questionnaire to scope the job, a day-by-day cutover runbook (the one canonical order for the whole toolkit), plain-English emails for staff and clients, a risk register, a handover template, and the list of what no copy tool moves. All Markdown, all meant to be copied and filled in for your own project.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.3.0 |
| Released | 10 November 2026 |
| Licence | MIT |

Part of my [M365 tenant-to-tenant migration toolkit](../README.md).

## Why these exist

The scripts move things. What sinks most migrations I've seen isn't a failed script, though. It's a question nobody asked at the start, a date nobody agreed, or staff who didn't know they'd need to sign in again on Monday. These are the documents I use to stop that happening.

## What happens without them

Scope creeps because nothing was written down. Cutover steps run out of order because nobody agreed the order. Staff and clients hear about the move from a bounce message. And a week later someone asks where the Planner board went, and the honest answer is that nobody said it wouldn't move.

## The documents

| Document | When | What it's for |
| --- | --- | --- |
| [Discovery questionnaire](Discovery-Questionnaire.md) | Before you quote | Questions about tenants, domains, DNS, licences, sync, devices, security, retention, mail flow, Teams and SharePoint, sizes, deadlines and who owns comms |
| [Risk register](Risk-Register.md) | From discovery to go/no-go | Common risks with likelihood, impact, mitigation and the toolkit part that helps |
| [Cutover runbook](Cutover-Runbook.md) | T-14 to T+7 | The canonical order. Day-by-day steps that call the toolkit scripts, the device move from T+0 to the source licence end, licensing timeline, coexistence, throttling, go/no-go checks, the rollback point, a worked rollback after the release, and who does what |
| [Comms templates](Comms-Templates.md) | T-7 to T+1 | Emails for staff, a notice for clients and suppliers, postponement and done-early notes, and updates for the business owner |
| [What doesn't move](What-Doesnt-Move.md) | Discovery, then T-7 | Teams chats, Planner, Forms, Power Platform, Loop, Bookings, phones, Intune and Windows devices, public folders, SSO apps, encrypted content and more: what goes wrong, who recreates it, three ways each |
| [Handover template](Handover-Template.md) | T+7 | What moved, where things are, admin accounts, DNS before and after, licences, open items, support window and decommission date |
| [DNS move playbook](DNS-Move-Playbook.md) | T-14, if the DNS host is out of reach | Moving DNS hosting to the registrar (or anywhere) before cutover: full scan, replicate, verify, switch name servers, rollback |
| [Lessons learnt](Lessons-Learnt.md) | Before you start | What went wrong, or nearly did, on a live cutover, and which script now handles each one |

## How to use them

1. Copy the `docs` folder into your project notes. Don't edit them here.
2. Start with the questionnaire. Its answers feed the risk register and the runbook dates.
3. Review the risk register with the business owner before every go/no-go.
4. Hand the comms templates to whoever owns staff communication, with dates filled in.
5. Fill in the handover as you go, not on the last day.

Keep filled-in copies out of public repos. They'll hold names, dates and DNS details.

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 1.3.0 | 2026-11-10 | Devices across the docs: runbook device section, Intune in what doesn't move, device lessons and risks 34 to 36 |
| 1.2.0 | 2026-11-10 | What doesn't move added. Runbook is the canonical order, with licensing, coexistence, throttling and a worked rollback. More comms templates. Discovery questions for what doesn't move |
| 1.1.0 | 2026-11-10 | DNS move playbook and lessons learnt added. Runbook and risk register updated after a live cutover |
| 1.0.0 | 2026-11-09 | First release |
