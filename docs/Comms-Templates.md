# Comms templates

**TL;DR:** Plain-English emails for staff at each stage of a tenant-to-tenant move, plus a short update for the business owner. No jargon: nobody outside IT needs to know what a tenant is. Swap the bits in [square brackets], cut anything that doesn't apply, and send them from someone people know, not from IT.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.0.0 |
| Released | 9 November 2026 |
| Licence | MIT |

Part of the [docs](README.md) for my [M365 tenant-to-tenant migration toolkit](../README.md). The timings match the [cutover runbook](Cutover-Runbook.md).

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
- Make sure anything in OneDrive or on your desktop is saved and synced.
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
