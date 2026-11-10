# Comms templates

**TL;DR:** Plain-English emails for staff at each stage of a tenant-to-tenant move, a notice for clients and suppliers, a postponement notice, an "it's done early" note, and updates for the business owner, including a one-message summary for cutover night. No jargon: nobody outside IT needs to know what a tenant is. Swap the bits in [square brackets], cut anything that doesn't apply, and send them from someone people know, not from IT.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.1.0 |
| Released | 10 November 2026 |
| Licence | MIT |

Part of the [docs](README.md) for my [M365 tenant-to-tenant migration toolkit](../README.md). The timings match the [cutover runbook](Cutover-Runbook.md).

## Why these exist

Most of the calls on the first morning aren't about a broken migration. They're from people who didn't know they'd need to sign in again, or clients whose email bounced for ten minutes and nobody warned them.

**Without them:** IT writes the emails at 11pm, in IT language, and sends them to the mailbox that's moving. People don't read them, and the helpdesk queue fills with questions the email should have answered.

## A few rules I follow

- **Send from a person.** The business owner or office manager, not a no-reply address. People read emails from people.
- **Say what they need to do, not what IT is doing.** "You'll need to sign in again on Monday" lands. "We're migrating your mailbox to a new tenant" doesn't.
- **Don't send passwords by email.** Especially not to the old mailbox, which is the thing that's changing.
- **Keep the old address in the message.** People panic when they think their email address is changing. Usually it isn't.
- **Print the cutover-morning email.** If someone can't sign in, they can't read an email telling them how to sign in.

## T-7: heads-up

**Subject:** Your email and files are moving on [date]. Nothing to do yet

Hi all,

On [day, date], we're moving our email, Teams and files to a new Microsoft 365 setup. [One line on why, e.g. "This is part of joining the Contoso group" or "This brings everything under one account".]

**What stays the same**
- Your email address stays as it is.
- Your email, calendar, contacts and files come with you.

**What will change**
- On [first working day], you'll sign in again on your laptop, in Outlook and Teams, and on your phone. We'll send instructions and be on hand to help.
- Teams chats from before the move won't come across. If there's anything in a chat you need to keep, save it somewhere else before [date]. [Files in Teams channels do come across. Check this against your copy option.]
- [If devices are in scope: Your laptop will need a short visit from [name] on [date]. It takes about [time].]

**What we need from you now**
- Nothing yet. Just keep [date] in mind.
- If you're on holiday or away around [date], let [name] know so we can plan for you.

Any questions, ask [name] at [contact].

Thanks,
[Name]

## T-1: what changes tomorrow

**Subject:** Tomorrow: signing in to email and Teams

Hi all,

The move happens [tonight / this weekend]. Here's what to expect when you start [tomorrow / on Monday].

**Before you finish today**
- Save and close everything you're working on.
- Make sure anything in OneDrive or on your desktop is saved and synced. Close OneNote once it has synced.
- [If MFA will be set up again: Have your work phone with you tomorrow. You'll need it to set up sign-in.]

**Tomorrow morning**
- You'll get a new temporary password from [name / your manager / in person at the office]. It won't be emailed.
- Your username is your normal email address.
- We'll send step-by-step instructions first thing, and there's a printed copy at [location].
- [Name] will be [in the office / on Teams / on the phone at number] from [time] to help.

**Please don't**
- Send emails to big groups or customers tonight. Some messages sent during the move may arrive late.
- Panic if Outlook asks you to sign in or says it's offline. That's expected.

Thanks for bearing with us,
[Name]

## Cutover morning: how to sign in

**Subject:** Signing in today: step by step

Good morning,

The move is done. Here's how to get back in. It takes about 10 minutes.

**1. Your laptop**
- Sign in to Windows as you normally do. [Adjust if devices were reset or rejoined.]

**2. Outlook and Teams**
- Open Outlook. If it asks you to sign in, use your email address and the temporary password you were given.
- You'll be asked to choose a new password. Pick something you haven't used before.
- [You'll be asked to set up the Microsoft Authenticator app. Follow the prompts on screen with your phone.]
- If Outlook still shows an error or old mail only, close it, wait a minute and open it again. If that doesn't work, tell [name].
- Open Teams and sign in the same way.

**3. Your phone**
- In Outlook on your phone, remove your work account, then add it again with your email address and new password.
- Do the same in Teams.
- [If people use the built-in mail app: remove the work account and add it again.]

**4. Check**
- Send yourself an email from your phone. Check it arrives in Outlook on your laptop.
- Open a Teams channel you use and check you can see the files.

Your older email might take a little while to appear in full. Files in OneDrive will follow over the next [day / few days]; we'll let you know when they're all there.

Stuck? [Name] is at [location / contact] until [time].

[Name]

## T+1: FAQ

**Subject:** Questions about the move

Hi all,

Thanks for your patience. Here are the questions we've been asked most.

**My old emails are missing.**
They're still coming across. Large mailboxes and archives take longer. If something's still missing on [date], tell [name].

**People say my replies are bouncing.**
Tell [name] who sent it and when. We can fix it, but we need an example.

**Outlook keeps asking for my password.**
Close Outlook completely and open it again. If it keeps happening, [name] can reset it.

**My shared mailbox isn't showing in Outlook.**
It can take a few hours to appear. If it's not there tomorrow, let [name] know which one.

**My files in OneDrive aren't there yet.**
They're being copied now and should be complete by [date].

**Where have my Teams chats gone?**
Chats from before the move don't come across. Channel conversations and files [do / don't, adjust to your copy option]. Start new chats as normal.

**A link to a shared file doesn't work.**
Links to files in the old setup won't work any more. Find the file in its new location and share it again.

**My phone stopped getting email.**
Remove your work account from the app and add it again. Instructions are in yesterday's email.

Anything else, ask [name].

[Name]

## T-7: notice for clients and suppliers

Only for the people outside the company who email you a lot: key clients, suppliers, your accountant. Send it from the business owner, a week ahead. Keep it short; they don't need the detail.

**Subject:** A short pause in our email on [day, date]

Hello [name],

On [day, date], between [time] and [time], we're moving our email to a new system. Our email addresses aren't changing.

During that window, a message you send us might bounce back or arrive late. If you get a bounce, please send it again after [time], or call us on [phone].

Anything urgent that evening, please phone [name] on [phone].

Thank you,
[Name]
[Company]

## Postponement notice

For staff, and for anyone outside who got the notice above. Send it as soon as the go/no-go is a no, or the night after a rollback. Short, calm, and with the new date if you have one.

**Subject:** The email move planned for [date] is postponed

Hi all,

We've decided not to go ahead with moving our email and files on [date]. [One honest line, e.g. "A check we ran beforehand found something we want to fix first." or "We started the move and chose to put things back as they were."]

**What this means for you**
- Nothing changes. Keep working exactly as you are, with the same sign-in and password.
- [If a rollback happened: If you signed in to anything new last night, ignore it. Use your usual sign-in.]
- [If anything was sent during a rollback: If you sent or expected an email between [time] and [time] last night, please check it arrived, and resend if not.]

**What happens next**
- The new date is [date]. / We'll confirm a new date by [date].
- You'll get the usual reminder the week before and the day before.

Thanks for your patience,
[Name]

## "It's done early" notice

When cutover went smoothly and finished well before the window you warned people about. Send it the same evening, or first thing, so people don't sit waiting.

**Subject:** The email move is done, ahead of schedule

Hi all,

Good news: the move finished early, at [time] [tonight / today]. Email is flowing normally.

**What to do**
- Nothing until [tomorrow morning / Monday]. Then follow the sign-in steps we'll send first thing.
- [If people can start now: If you want to sign in now, use your email address and the temporary password you'll get from [name].]

[For clients, if you sent them the notice: "The pause in our email we mentioned for [date] is over, earlier than expected. Everything is working normally."]

Thanks,
[Name]

## Business owner update

Short, non-technical, and honest. Send it at T-7, after cutover, and at sign-off.

**Subject:** Migration update: [date]

Hi [name],

**Where we are:** [One sentence. e.g. "Cutover went ahead on Friday evening as planned."]

**What's done**
- [e.g. Everyone can sign in and email is flowing in and out.]
- [e.g. Shared mailboxes and team addresses are working.]

**What's still in progress**
- [e.g. Older email and archives for three people, due to finish by Wednesday.]
- [e.g. OneDrive files, due by Thursday.]

**Problems and what we're doing about them**
- [e.g. Two phones needed help reconnecting. Both fixed.]
- [Or: "None so far."]

**Decisions needed from you**
- [e.g. The old system can be switched off on [date]. OK to go ahead?]
- [Or: "None this week."]

**Next update:** [date]

[Name]

## Cutover night: one-message summary for the business owner

One message, the night it happens or first thing next morning. It answers "is it done, and what do I need to know?" without making them open anything. Fill it from the 09 summary, the migration report, and the sign-in check. No names of individuals, no passwords, no technical detail they'd have to ask about.

**Subject:** Done: email and files moved [tonight / on date]

Hi [name],

**Result:** The move is done. Email to and from [fabrikam.com] is flowing in the new system, and test sign-ins worked.

**Timings**
- Started [time], domain moved at [time], mail flowing again at [time]. Total time email could bounce: about [x] minutes.

**Checks**
- Validation: [x] of [y] accounts fully checked and passing. [z] with a warning we're looking at, none blocking. / [z] need a fix tomorrow, listed below.
- Sign-in: test sign-ins for [2] people worked from [location]. Everyone will sign in with a temporary password tomorrow morning.
- Mail: test messages in and out from an outside address arrived.

**Still to finish**
- [e.g. Older mail and archives for a few large mailboxes: due by [date].]
- [e.g. OneDrive files: copying from tomorrow, due by [date].]

**What happens tomorrow**
- [Name] hands out temporary passwords from [time]. [Name] is on hand at [location] until [time].
- Old accounts are switched off [tomorrow / on date], so nobody keeps working in the old system.

**Anything you need to decide:** [e.g. "Nothing tonight." / "OK to switch off the old accounts tomorrow at noon?"]

I'll send a fuller update at [time / date].

[Name]

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 1.1.0 | 2026-11-10 | Client and supplier notice, postponement notice, "done early" notice, cutover night one-message summary for the business owner |
| 1.0.0 | 2026-11-09 | First release |
