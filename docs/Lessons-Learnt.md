# Lessons learnt from a live cutover

**TL;DR:** What went wrong, or nearly did, on a real small-business tenant-to-tenant cutover, and what changed in the toolkit because of it. Grouped into before the night, on the night, and after. Each lesson says what happened, what it cost, the fix, and which script now handles it. Names, domains and tenants are removed; examples use fabrikam.com moving into the contoso tenant.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.0.1 |
| Released | 10 November 2026 |
| Licence | MIT |

Part of the [docs](README.md) for my [M365 tenant-to-tenant migration toolkit](../README.md).

## The short list

If you read nothing else:

1. Find out who can log in to the **DNS host**, not just the registrar, two weeks out.
2. Add the domain to the destination **first**, publish its TXT, then release it from the source.
3. The day before, take the `-pre` snapshot, then run `07 Test-CutoverReadiness.ps1` with it. Do both again on the night before the release. It finds the blockers you'd otherwise meet at 9pm. The order is in the [cutover runbook](Cutover-Runbook.md).
4. Every script that changes a tenant takes `-AdminUpn`. Use it.
5. One version of every Microsoft.Graph module, and a fresh PowerShell window per task.
6. After cutover, rebuild the user mapping before any OneDrive copy or incremental.

## Before the night

### The DNS host wasn't the registrar

| | |
| --- | --- |
| What happened | The domain was registered at one company, but its name servers pointed at a cloud DNS service in an account nobody on the project could log in to. The registrar login worked; the DNS records weren't there |
| Impact | No way to add the verification TXT or change MX on the night. Cutover would have stalled after the source released the domain, with mail bouncing |
| Fix | Scanned every record from public DNS, recreated them at the registrar's own DNS through its API, checked them against the new name servers, then switched NS, two days before cutover |
| Now handled by | [07 Export-DomainDns.ps1 -FullScan](../07-DomainCutover/), [07 Set-DomainDnsRecords.ps1](../07-DomainCutover/), and the [DNS move playbook](DNS-Move-Playbook.md) |

**Tip:** `Resolve-DnsName fabrikam.com -Type NS` tells you where DNS really lives. If the answer isn't the company whose login you have, start the move now.

### Domain removal blockers turned up late

| | |
| --- | --- |
| What happened | `Remove-SourceDomain -WhatIf` listed deleted users still in the recycle bin, a user who existed only as a proxy address on another object, and an enterprise app with a reply URL on the domain. Changes made in Exchange took a while to show in Entra ID (sync lag), so a second run still showed blockers that had been fixed |
| Impact | Each blocker would have stopped the removal on the night |
| Fix | Cleared them the day before: deleted users renamed and re-deleted, proxy addresses removed, the app's URL changed. Waited, re-ran, confirmed empty |
| Now handled by | [07 Remove-SourceDomain.ps1 -BlockersOnly](../07-DomainCutover/) at T-7 and T-1, [07 Clear-DeletedUserDomain.ps1](../07-DomainCutover/), [07 Test-CutoverReadiness.ps1](../07-DomainCutover/) |

**Trap:** a clean `-WhatIf` an hour after a fix isn't proof. Exchange and Entra ID sync on their own schedule. Re-run the next morning.

### Snapshot timing

| | |
| --- | --- |
| What happened | The address snapshot (every recipient's addresses and LegacyExchangeDN) was first taken days early. People and lists changed after it |
| Impact | X500s and old aliases for anything created or changed after the snapshot would have been missed, and replies to old mail would bounce |
| Fix | Took the snapshot again on the night, immediately before `-RemoveDomain`. That's the one that counts |
| Now handled by | `Remove-SourceDomain.ps1` snapshots again before removing; [07 Test-CutoverReadiness.ps1](../07-DomainCutover/) checks the snapshot file is there and is the right one |

### Guests in the destination held the old addresses

| | |
| --- | --- |
| What happened | The destination had guest accounts for source staff, from earlier collaboration. Each guest carried the person's fabrikam.com address |
| Impact | Adding that address to the real mailbox at cutover failed: "address already in use". Mapping tools also matched people to their own guest |
| Fix | Removed or renamed the guests before cutover, after checking nothing important was shared to them |
| Now handled by | [02b New-UserMapping.ps1](../02b-UserMapping/) never matches a guest; [07 Test-CutoverReadiness.ps1](../07-DomainCutover/) lists destination objects already holding source addresses; `Switch-TargetAddresses` reports them as Conflict |

### The destination hosted other companies

| | |
| --- | --- |
| What happened | The destination tenant already held several companies, each with its own `info@`, `accounts@` and `admin@` |
| Impact | Alias and display-name matching picked another company's shared mailbox |
| Fix | Scoped matching to the domains this migration owns, and prefixed staged objects |
| Now handled by | `-DestinationDomain` and `-AliasPrefix` on [02b](../02b-UserMapping/); `-AliasPrefix` on [05](../05-MailConfig/) and [06](../06-GroupsTeamsSites/) |

### Graph module version clash

| | |
| --- | --- |
| What happened | `Microsoft.Graph.Authentication` 2.41.2 and `Microsoft.Graph.Identity.DirectoryManagement` 2.26.1 were both installed. A script failed with "Could not load file or assembly 'Microsoft.Graph.Authentication, Version=2.26.1.0' ... Assembly with same name is already loaded" |
| Impact | A destination script stopped part way. Fixing it in the same window didn't help |
| Fix | Installed every Microsoft.Graph module at one version, closed every window, opened a fresh one |
| Now handled by | [tools/Test-ModuleHealth.ps1](../tools/) (and `-Fix`), which `Test-CutoverReadiness` runs too |

### Scripts reused the wrong tenant's session

| | |
| --- | --- |
| What happened | With both tenants open on one PC, a script picked up the Exchange Online session left open from the other tenant |
| Impact | The tenant ID guard stopped it, but only after a confusing error, and time lost on the night |
| Fix | Disconnect between tenants, and name the account you mean |
| Now handled by | `-AdminUpn` on [02](../02-TargetMailboxes/), [04](../04-TargetOneDrives/), [05](../05-MailConfig/), [06](../06-GroupsTeamsSites/), [09](../09-Validation/) and the 07 scripts. If the open session is anyone else, the script disconnects, signs in as that account and prints "Signed in as X to tenant Y" before it changes anything |

### ShareGate's PowerShell only loads in 5.1

| | |
| --- | --- |
| What happened | The ShareGate module failed to load in PowerShell 7 |
| Impact | The bulk OneDrive copy script couldn't run from the window the other scripts used |
| Fix | Two windows: `pwsh` (7) for the toolkit, `powershell.exe` (5.1) for ShareGate |
| Now handled by | `Copy-DrivesWithShareGate.ps1` stops with the right command in 7; [Test-ModuleHealth](../tools/) reports it |

## On the night

### Add first, then release, then verify

| | |
| --- | --- |
| What happened | The runbook said release from the source, then add to the destination. Adding first works: the destination gives you the TXT record even while the source still holds the domain. Only verification has to wait |
| Impact | Releasing first adds the TXT publishing time (and DNS propagation) to the window where mail can bounce |
| Fix | Add to the destination, publish the TXT and wait for it to resolve, then release from the source, then verify straight away |
| Now handled by | [07 Add-TargetDomain.ps1 -WaitForDns](../07-DomainCutover/); the [cutover runbook](Cutover-Runbook.md) T-0 order |

### MX and SPF didn't need to change

| | |
| --- | --- |
| What happened | Both tenants are Microsoft 365, and the MX and SPF records the destination asked for were the same as the ones already published |
| Impact | None, once we realised. Mail follows the domain to whichever tenant has it verified. The real gap was between removal and verification, not DNS |
| Fix | Compared what `Add-TargetDomain -Verify` printed with the live records and changed only what differed |
| Now handled by | `Add-TargetDomain.ps1` compares each record with live DNS and marks it unchanged |

**Trap:** newer domains can be given a different MX format. Always read what the destination asks for; don't assume it's the same.

### A dangerous apex CNAME suggestion

| | |
| --- | --- |
| What happened | A record list suggested a CNAME on the root of the domain |
| Impact | A CNAME at the apex can't coexist with any other record. Publishing it would have broken MX, SPF, the verification TXT and the website |
| Fix | Didn't publish it |
| Now handled by | `Add-TargetDomain.ps1` marks any apex CNAME as Do not apply |

### -WhatIf suppressed the reports

| | |
| --- | --- |
| What happened | Under `-WhatIf`, PowerShell also skipped `Export-Csv` and `New-Item`, so the dry run left no report |
| Impact | Nothing to review or show the business owner before the real run |
| Fix | Every report write passes `-WhatIf:$false`, so a dry run still writes its CSV with WhatIf rows |
| Now handled by | All scripts; [tools/Get-MigrationReport.ps1](../tools/) reads those reports |

### Conditional Access blocked the admin's test sign-in

| | |
| --- | --- |
| What happened | The destination had a "block outside allowed countries" Conditional Access policy. The admin testing the moved accounts was in another country. Sign-in failed with error 53003 |
| Impact | Looked like a broken migration at the worst moment |
| Fix | Tested from inside the allowed location (a colleague on site). Left the policy alone |
| Now handled by | [09 Test-TargetTenant.ps1](../09-Validation/) warns about every policy with a location condition; [07 Test-CutoverReadiness.ps1](../07-DomainCutover/); the runbook's T-1 test sign-in plan |

**Tip:** if you must test from elsewhere, add your test account to the policy's exclusions for the night, write it down, and remove it in the morning.

### Blocking source sign-in

| | |
| --- | --- |
| What happened | Moved users could still sign in to the source with their old onmicrosoft.com name. Blocking them by script worked for everyone except accounts holding admin roles |
| Impact | People could have kept working in the old tenant |
| Fix | Blocked everyone else by script; blocked the admin-role accounts in the admin centre, last, keeping one break-glass account |
| Now handled by | [07 Block-SourceSignIn.ps1](../07-DomainCutover/) (reports AdminRole rows instead of failing); [10 Decommission](../10-Decommission/) |

## After

### The user mapping had to be rebuilt

| | |
| --- | --- |
| What happened | Sign-in names changed at cutover, so the mapping CSV and the ShareGate `.sgum` pointed at names that no longer existed |
| Impact | An incremental or OneDrive copy with the old files maps nobody: authors become the migration account and permissions go missing |
| Fix | Re-ran 02b after cutover, re-imported the `.sgum`, rebuilt the OneDrive copy list |
| Now handled by | [02b New-UserMapping.ps1](../02b-UserMapping/), [copy-options/New-DriveCopyList.ps1](../copy-options/) |

### Former staff OneDrives went into one archive library

| | |
| --- | --- |
| What happened | Leavers' OneDrives were copied into one library on an archive site, a folder per person, from a CSV built by hand. The folders were created by hand too |
| Impact | An evening of typing URLs, and two typos |
| Fix | Scripted the list and the folders |
| Now handled by | [copy-options/New-DriveCopyList.ps1](../copy-options/) `-Leavers -CreateFolders`, consumed by option B and option C `-CsvPath` |

### Validation gaps

| | |
| --- | --- |
| What happened | Validation didn't check that moved users must change their temporary password, and said nothing about the destination's sign-in policies. The destination had no Conditional Access policy requiring MFA |
| Impact | A temporary password could stay valid indefinitely; a security gap went unreported |
| Fix | Added the checks. The MFA gap is reported as a finding for the business owner; the toolkit doesn't change security policy |
| Now handled by | [09 Test-TargetTenant.ps1](../09-Validation/) Password, CA-Location and MFA checks |

### Leaving the old transport rules behind was right

| | |
| --- | --- |
| What happened | The source had years of transport rules: disclaimers for a brand that no longer existed, routing to a retired system, rules naming people who had left |
| Impact | Rebuilding them all would have imported old problems into a tenant that serves other companies too |
| Fix | Exported them as reference only, rebuilt the one that was still needed, disabled until tested |
| Now handled by | [05 Copy-MailConfig.ps1](../05-MailConfig/) exports rules, never imports them |

**Tip:** after an acquisition, treat the old company's mail flow rules as history, not configuration. The business owner decides which, if any, still matter. The same goes for forwarding to outside addresses.

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 1.0.1 | 2026-11-10 | Readiness and blocker steps match the cutover runbook. Job details generalised |
| 1.0.0 | 2026-11-10 | First version, written the morning after cutover |
