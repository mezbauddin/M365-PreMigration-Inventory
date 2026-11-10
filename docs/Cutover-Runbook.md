# Cutover runbook

**TL;DR:** The two weeks either side of cutover, day by day, with the exact toolkit step for each. This is the **one canonical order** for the whole toolkit: the [07 Domain cutover](../07-DomainCutover/) README and the root README follow it, and if they ever disagree, this page wins. It has the licensing timeline, a coexistence plan if you move in stages, throttling guidance for the big copies, go/no-go checks, a rollback point, a worked rollback procedure for after the domain has moved, and who does what. Copy it, put your dates and names in, and agree it with the business owner before T-14. Examples use fabrikam.com moving into the contoso tenant.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.2.0 |
| Released | 10 November 2026 |
| Licence | MIT |

Part of the [docs](README.md) for my [M365 tenant-to-tenant migration toolkit](../README.md).

## Why a runbook

Cutover night is a relay between two tenants and a DNS host, with mail bouncing in the gap. Every step depends on the one before: the snapshot must exist before the readiness check can read it, the name servers must move two days before anything else can change, the domain must be verified before the addresses can switch. Do them out of order and you find out at 9pm.

**Without it:** someone runs the release before the snapshot, or moves DNS on the night, or lets the source licences lapse before the last incremental. Each of those is recoverable. None of them is fun.

## Who does what

| Role | Who | Does |
| --- | --- | --- |
| Migration lead | [name] | Runs the scripts and copy jobs. Owns this runbook |
| Business owner | [name] | Makes the go/no-go call. Signs off at T+7 |
| DNS owner | [name, phone] | Can log in to the **DNS host** (where the name servers point), not just the registrar. Changes DNS records on the night. Must be reachable during the cutover window |
| Comms owner | [name] | Sends the staff and external emails in [Comms templates](Comms-Templates.md) |
| Licensing owner | [name] | Buys destination licences and add-ons, and controls the source renewal date |
| Day-one support | [names] | On hand on the first working morning for sign-ins, phones and Outlook |
| Device support | [names] | Rejoins or resets laptops and phones, if that's in scope ([08](../08-Devices/)) |

Swap names in. One person can hold several roles in a small business, but the DNS owner and the business owner must be named and reachable on the night.

## The order at a glance

| When | Step | Part |
| --- | --- | --- |
| T-14 | Inventory, build staged mailboxes, mapping, EWS fix, groups, first copy | 01, 02, 02b, 03, 06, copy option |
| T-14 | DNS full scan. Any DNS host or name server move happens **now** (NS TTL is often 48 hours) | 07 `Export-DomainDns -FullScan`, `Set-DomainDnsRecords` |
| T-14 | Module health on the cutover PC | tools |
| T-7 | Mail config, SharePoint and Teams copies, licence shared mailboxes | 05, 02 `-SharedMailboxesOnly` |
| T-7 | List and clear blockers | 07 `Remove-SourceDomain -BlockersOnly`, `Clear-DeletedUserDomain` |
| T-7 | Lower TTLs, first validation, heads-up emails | 07, 09, comms |
| T-3 | Incremental, fresh inventory, validation, dry run | copy option, 01, 09 |
| T-1 day | `-pre` snapshot, blockers again, readiness check, go/no-go | 07 `Remove-SourceDomain -SnapshotOnly`, `-BlockersOnly`, `Test-CutoverReadiness` |
| T-0 | Add domain to destination, TXT, night's `-pre` snapshot, final readiness run | 07 |
| T-0 | **Rollback point** | |
| T-0 | Release, verify, DNS, license and convert, switch addresses, OneDrives, test, validate, DKIM | 07, 02 `-Apply`, 04, 09 |
| T+1 | Block source sign-in, rebuild mapping, OneDrive copy | 07, 02b, copy option |
| T+7 | Sign-off, decommission date | 09, [10](../10-Decommission/) |

## Licensing timeline

Licences are the step people forget to put dates on. Get these four dates agreed in writing before T-14.

| Date | What | Typical timing |
| --- | --- | --- |
| Destination licences bought | Seats for everyone in `TargetBuildPlan.csv`, visible in the destination | Before T-14. Staged mailboxes don't need them, but shared mailboxes over 50 GB or with archives do, at T-7 |
| Migration add-on bought | If you use Microsoft's native cross-tenant migration, the per-user Cross-Tenant User Data Migration add-on. Check Microsoft's current guidance on which tenant holds it | At least a week before the first batch, so it shows up and can be assigned |
| Overlap month | Both tenants paid for at the same time | From the first copy until at least T+30. Budget a full month of double licensing; it is cheaper than a rushed last incremental |
| Source renewal or end date | When the source subscription renews, or stops | Never before T+30. Ideally move it to after the decommission date in [10](../10-Decommission/) |

**Trap:** don't let the source lapse, or drop its licences, before the final incrementals. When a licence is removed, the mailbox and OneDrive go into retention and the copy tool can't read them. On a monthly subscription that's one missed renewal away.
**Trap:** removing a licence from a source user with an archive or a hold has its own consequences. Decide per mailbox in [10](../10-Decommission/), not in a bulk licence removal.
**Tip:** if the source is on an annual commitment, find the renewal date in discovery. It often decides the cutover date for you.

## Coexistence if you move in stages

Most small businesses move everyone in one night, and this runbook assumes that. If you move in waves, people in the two tenants have to work together for weeks. Plan it, or they'll email each other's old addresses and get nothing.

| Need | How | Notes |
| --- | --- | --- |
| Mail for movers reaches the new mailbox | Forward each moved source mailbox to its destination onmicrosoft.com address (`Set-Mailbox -ForwardingSmtpAddress`, `-DeliverToMailboxAndForward $false`) | The source outbound spam policy blocks automatic external forwarding by default. Allow it for the movers only, with a custom outbound policy scoped to them |
| Free/busy across tenants | An organisation relationship in each tenant (`Get-FederationInformation`, then `New-OrganizationRelationship` with free/busy access) | Gives availability, not full calendar sharing. Test with a meeting request both ways |
| Find each other in the address book | Mail contacts in each tenant for the people in the other | [05](../05-MailConfig/) exports and rebuilds contacts; build the cross-tenant ones from the mapping |
| Teams chat between waves | External access (federation) allowed between the two tenants, both ways | Chats become external chats. Shared channels need B2B direct connect, which is more set-up than it's worth for a few weeks |
| Shared mailboxes | Move a shared mailbox with the people who use it most | Full Access across tenants doesn't work. That's the usual reason waves get merged |

**Trap:** forwarding from the source stops working the moment the domain moves, because the source no longer accepts mail for it. Remove the forwards in the same window as the release, or they'll error.
**Tip:** if there's no strong reason to move in waves, don't. One night of disruption is easier to explain than six weeks of coexistence.

Three ways:

- **Paid tool:** some migration suites run coexistence for you (directory sync between tenants, free/busy, mail routing). Worth it above a few hundred people.
- **Microsoft native:** organisation relationships, external access and, for larger organisations, cross-tenant synchronisation or a multi-tenant organisation.
- **By hand:** forwards, contacts and an organisation relationship, as above. Fine for a small business.

## Throttling and when to copy

Every copy route is throttled by Microsoft, and the limits are per tenant and per app. You can't buy your way past them; you can only work with them.

| Service | What you'll see | What to do |
| --- | --- | --- |
| Microsoft Graph (OneDrive and SharePoint copies by Graph, user and group reads) | HTTP 429 with a `Retry-After` header | Honour `Retry-After`. The toolkit's Graph copy already does. Running more parallel jobs makes it worse, not faster |
| EWS (ShareGate and other tools that copy mail and archives over EWS) | Slow items, "server busy" errors, jobs that crawl in the afternoon | Copy outside business hours in the tenant's region. Ask Microsoft support, or the admin centre's self-help diagnostics where offered, for a temporary EWS throttling increase for migration. EWS ends on 1 April 2027 |
| SharePoint and OneDrive (all tools) | 429 or 503, with `Retry-After` | Run big copies in the evening and at weekends. Microsoft throttles harder during the working day |
| Native cross-tenant mailbox moves | Batches that sit at "syncing" for longer than expected | These run on Microsoft's own move service, not EWS. Start early, and expect the first pass of a big archive to take days |

Batch sizes I start with, then adjust from what the reports show:

- **Mail and archives:** 10 to 20 mailboxes per batch, biggest archives in their own batch, started first.
- **OneDrive:** 5 to 10 at a time. One person with 300 GB will take longer than everyone else together.
- **SharePoint and Teams:** one big site at a time, small ones in parallel.

**Tip:** start the first full pass at T-14 on a Friday evening. By Monday you'll know your real throughput, and whether T-0 is realistic.
**Trap:** a re-run started while the first one is still throttled doubles the load and can make duplicates. Let a job finish or stop it cleanly before starting again.

## Before T-14

These should already be done. If they aren't, the dates move.

- [ ] Discovery done and agreed ([Discovery questionnaire](Discovery-Questionnaire.md)), including [What doesn't move](What-Doesnt-Move.md) and who rebuilds each item
- [ ] Copy option chosen ([copy-options](../copy-options/))
- [ ] Licensing timeline above agreed: destination licences, add-on, overlap month, source end date
- [ ] [Risk register](Risk-Register.md) reviewed with the business owner

## T-14: build, DNS and first copy

| Step | Who | Action |
| --- | --- | --- |
| 1 | Lead | Run [01](../01-SourceInventory/) against the source. Keep the folder |
| 2 | Lead | Run [02](../02-TargetMailboxes/) `-Plan`, review `TargetBuildPlan.csv` with the business owner, then `-Apply -StageOnly` against the destination |
| 3 | Lead | Run [02b](../02b-UserMapping/) to build the mapping file and `.sgum`. Check the review list line by line |
| 4 | Lead | If your tool copies archives over EWS, run [03](../03-EwsArchiveReadiness/) with `-Fix` in both tenants |
| 5 | Lead | Run [06](../06-GroupsTeamsSites/) `-Export` then `-Import` so groups and Teams exist before you copy into them |
| 6 | Lead | Import the user mapping into the copy tool. Start the first mail and archive copy, outside hours, in batches (see throttling) |
| 7 | Lead | Run [07](../07-DomainCutover/) `Export-DomainDns.ps1 -FullScan`. It prints where DNS really lives and the NS TTL. Send the DNS owner the list of TTLs to lower |
| 8 | Lead and DNS owner | Confirm who can log in to the **DNS host**. If the name servers aren't at a company whose login you have, move DNS hosting **now** with `Set-DomainDnsRecords.ps1 -FromScan`, then switch the name servers. The NS TTL is often 48 hours, so this can't wait until T-3 ([DNS move playbook](DNS-Move-Playbook.md)) |
| 9 | Lead | Run [tools/Test-ModuleHealth.ps1](../tools/) on the PC you'll use on the night. Fix any Graph version mismatch now, not on the night |
| 10 | Comms | Agree the cutover date in writing with the business owner |

**Trap:** the registrar and the DNS host are often different companies. Having the registrar login proves nothing about who can add a TXT record on the night.

```powershell
.\Invoke-SourceTenantInventory.ps1 -SharePointAdminUrl https://fabrikam-admin.sharepoint.com
.\New-TargetMailboxes.ps1 -Plan -TargetDomain contoso.onmicrosoft.com
.\New-TargetMailboxes.ps1 -Apply -TargetDomain contoso.onmicrosoft.com -UsageLocation GB -StageOnly -AdminUpn admin@contoso.onmicrosoft.com
.\New-UserMapping.ps1 -SourceAdmin admin@fabrikam.onmicrosoft.com -DestinationAdmin admin@contoso.onmicrosoft.com
.\Export-DomainDns.ps1 -Domain fabrikam.com -FullScan
.\Test-ModuleHealth.ps1
# Only if the DNS host is out of reach: replicate the zone somewhere you control, then switch NS (see the DNS move playbook)
.\Set-DomainDnsRecords.ps1 -Domain fabrikam.com -Provider <GoDaddy or Cloudflare> -FromScan .\DomainDns-FullScan-<stamp>.csv -WhatIf
```

## T-7: config, blockers and heads-up

| Step | Who | Action |
| --- | --- | --- |
| 1 | Lead | Check the first copy finished. Fix and re-run anything that failed |
| 2 | Lead | Run [02](../02-TargetMailboxes/) `-Apply -SharedMailboxesOnly` now that licences are in: shared mailboxes that need a licence or archive get one |
| 3 | Lead | Run [05](../05-MailConfig/) `-Export` against the source, then `-Import -WhatIf` and `-Import` against the destination |
| 4 | Lead | Start SharePoint and Teams copies into the groups 06 created |
| 5 | Lead | Run [07](../07-DomainCutover/) `Remove-SourceDomain.ps1 -BlockersOnly`. It's read-only. Clear every blocker it lists: `Clear-DeletedUserDomain.ps1` for deleted users, the admin centre Remove domain wizard for ProxyOnly users, sync off for synced users. Application blockers are fixed on the night by `-FixApplications`; tell the app owners now |
| 6 | Lead | Look for guests and other objects in the destination already holding source addresses. `Test-CutoverReadiness` lists them at T-1; decide now whether `Switch-TargetAddresses -RemoveConflictingGuests` deletes them on the night or you remove them by hand |
| 7 | Lead | Run [09](../09-Validation/) without `-Domain`. Expect warnings; read every fail |
| 8 | Comms | Send the T-7 heads-up to staff and the external senders notice to clients and suppliers |
| 9 | DNS owner | Lower TTLs on MX, autodiscover, SPF, DKIM and any CNAMEs flagged |

```powershell
.\Remove-SourceDomain.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -Domain fabrikam.com -BlockersOnly
.\Clear-DeletedUserDomain.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -Domain fabrikam.com -WhatIf
```

## T-3: top up and dry run

- [ ] Incremental copy of mail and archives
- [ ] Re-run 01 so item counts for validation are recent
- [ ] Re-run 09 and compare with T-7. Fails should be going down
- [ ] `Export-DomainDns.ps1` (default run) to confirm the TTLs really are lower
- [ ] Walk through the cutover night steps with the DNS owner, out loud, with times
- [ ] Confirm day-one support people and where they'll be
- [ ] Prepare the temporary password hand-out method (in person, by phone, or by manager). Not by email to the old mailbox

## T-1 day: snapshot, readiness and go/no-go

The order matters: the readiness check reads the snapshot, so the snapshot comes first.

1. `Remove-SourceDomain.ps1 -SnapshotOnly`. Writes `SourceRecipients-pre-<stamp>.csv`, every recipient with its addresses and LegacyExchangeDN while the source still holds the domain. Keep it safe. If the night goes badly, this is your fallback copy.
2. `Remove-SourceDomain.ps1 -BlockersOnly` again. It should show nothing but Application blockers (fixed on the night by `-FixApplications`) and Info lines.
3. `Test-CutoverReadiness.ps1` with `-SourceSnapshot` pointing at the file from step 1. It runs the module check, reads both tenants and DNS, and gives PASS, WARN and FAIL.
4. `Get-MigrationReport.ps1 -Html` for every report so far, on one page.

```powershell
.\Remove-SourceDomain.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -Domain fabrikam.com -SnapshotOnly
.\Remove-SourceDomain.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -Domain fabrikam.com -BlockersOnly
.\Test-CutoverReadiness.ps1 -SourceTenantId <source-guid> -DestinationTenantId <dest-guid> `
    -SourceAdmin admin@fabrikam.onmicrosoft.com -DestinationAdmin admin@contoso.onmicrosoft.com `
    -Domain fabrikam.com -MappingFile .\UserMapping.csv -SourceSnapshot .\SourceRecipients-pre-<stamp>.csv
.\Get-MigrationReport.ps1 -Html
```

The business owner makes the call. Don't start cutover with a no.

| Check | Go if |
| --- | --- |
| Readiness | `Test-CutoverReadiness.ps1` has no FAIL. Every WARN read and accepted |
| Copy status | Every mailbox has had at least one complete pass. Archives done, or agreed to finish after cutover |
| Validation | 09 has no unexplained fails. Every remaining fail has an owner and a plan |
| Licences | Enough seats for everyone in the plan. Source paid up past T+30 |
| Blockers | `Remove-SourceDomain -BlockersOnly` shows none left, apart from Application blockers `-FixApplications` will handle |
| DNS | TTLs were lowered at least a day ago. DNS owner can log in to the DNS host (tested, not assumed) and is confirmed for the night. Any name server move finished at least 48 hours ago |
| Test sign-in plan | Who tests the moved accounts, from where. If the destination has location-based Conditional Access, the tester must be in an allowed location, or their test account excluded for the night and the exclusion removed after. Otherwise sign-in fails with error 53003 and looks like a broken migration |
| People | Day-one support confirmed. Comms owner ready to send the T-1 email |
| Business | No month end, payroll run or key event tomorrow |

If it's a go:

- [ ] Comms sends the T-1 email
- [ ] Lead records the go decision, time and who made it

If it's a no: comms sends the postponement notice from [Comms templates](Comms-Templates.md) to staff and to the external senders who were told.

## T-0: cutover night

Do these in order. Note the time against each step. This is the **add-first** order: the destination gives you its verification TXT while the source still holds the domain, so the TXT is published and resolving before anything changes for users. Only verification has to wait for the source to let go.

| # | Time | Who | Action |
| --- | --- | --- | --- |
| 1 | | Lead | Final incremental copy of mail. Let it finish |
| 2 | | Lead | `Export-DomainDns.ps1` again. This file is your rollback copy of DNS |
| 3 | | Lead | Destination: `Add-TargetDomain.ps1`. Send the TXT record it prints to the DNS owner |
| 4 | | DNS owner | Add the TXT record. Leave everything else alone, **including the source's old MS=ms TXT**: it is what lets you roll back |
| 5 | | Lead | `Add-TargetDomain.ps1 -WaitForDns` until every authoritative name server serves the TXT |
| 6 | | Lead | Source: `Remove-SourceDomain.ps1 -SnapshotOnly`. This is the night's `SourceRecipients-pre-<stamp>.csv`, the one Switch-TargetAddresses and 09 use. It catches anything that changed since T-1 |
| 7 | | Lead | Final `Test-CutoverReadiness.ps1` with the night's `-pre` file. No FAIL, or stop |
| | | **Lead and business owner** | **Rollback point. Last chance to stop cheaply. Before this line, nothing has changed for users. The TXT record does no harm if you stop** |
| 8 | | Lead | `Remove-SourceDomain.ps1 -FixApplications -RemoveDomain`. App identifier URIs move to onmicrosoft.com, then the domain leaves the source. Mail to it bounces until step 9. (This run writes one more `-pre` snapshot before it strips anything; either that or step 6's file will do) |
| 9 | | Lead | Disconnect from the source. `Add-TargetDomain.ps1 -Verify`. If it says the domain is still in use, wait a few minutes and retry. Send the records it marks Change or Add to the DNS owner. Records marked Unchanged stay as they are, and never apply anything marked DoNotApply |
| 10 | | DNS owner | Change only what step 9 asked for: often just the DKIM CNAMEs. MX, SPF and autodiscover are usually the same between two Microsoft 365 tenants |
| 11 | | Lead | [02](../02-TargetMailboxes/) `New-TargetMailboxes.ps1 -Apply` (no `-StageOnly`): staged people become licensed user mailboxes with archives. Run twice, about 15 minutes apart |
| 12 | | Lead | `Switch-TargetAddresses.ps1 -WhatIf` with the night's `-pre` file, read every Conflict, then run with `-RemoveConflictingGuests -ResetPasswords` (leave out `-RemoveConflictingGuests` if you decided at T-7 to deal with guests by hand). These passwords replace any step 11 set, so hand out the ones from this file, with the new sign-in names |
| 13 | | Lead | [04](../04-TargetOneDrives/) `Request-TargetOneDrives.ps1` now that people are licensed. `-Wait` if you have time; otherwise check its report in the morning |
| 14 | | Lead | Send a test message in and out from an outside address. Check it lands |
| 15 | | Tester | Test sign-in for two moved users, from an allowed location (see the T-1 test sign-in plan) |
| 16 | | Lead | `Test-TargetTenant.ps1` ([09](../09-Validation/)) with `-Domain` and `-SourceSnapshot` |
| 17 | | Lead | `Add-TargetDomain.ps1 -EnableDkim` once the CNAMEs resolve. Can wait until morning |
| 18 | | Lead | Final incremental copy of mail, now that addresses are final |
| 19 | | Lead | `Get-MigrationReport.ps1 -Html`. Send the business owner the one-message summary from [Comms templates](Comms-Templates.md), with the 09 summary and the report |

```powershell
# Every tenant-changing script takes -AdminUpn: it disconnects any other session and prints who and where before changing anything
.\Export-DomainDns.ps1 -Domain fabrikam.com -Server 1.1.1.1
.\Add-TargetDomain.ps1 -TenantId <dest-guid> -AdminUpn admin@contoso.onmicrosoft.com -Domain fabrikam.com
.\Add-TargetDomain.ps1 -TenantId <dest-guid> -AdminUpn admin@contoso.onmicrosoft.com -Domain fabrikam.com -WaitForDns
.\Remove-SourceDomain.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -Domain fabrikam.com -SnapshotOnly
.\Test-CutoverReadiness.ps1 -SourceTenantId <source-guid> -DestinationTenantId <dest-guid> `
    -SourceAdmin admin@fabrikam.onmicrosoft.com -DestinationAdmin admin@contoso.onmicrosoft.com `
    -Domain fabrikam.com -MappingFile .\UserMapping.csv -SourceSnapshot .\SourceRecipients-pre-<stamp>.csv
# ---- rollback point ----
.\Remove-SourceDomain.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -Domain fabrikam.com -FixApplications -RemoveDomain
Disconnect-ExchangeOnline -Confirm:$false; Disconnect-MgGraph
.\Add-TargetDomain.ps1 -TenantId <dest-guid> -AdminUpn admin@contoso.onmicrosoft.com -Domain fabrikam.com -WaitForDns -Verify
.\New-TargetMailboxes.ps1 -Apply -TargetDomain contoso.onmicrosoft.com -UsageLocation GB -AdminUpn admin@contoso.onmicrosoft.com
.\Switch-TargetAddresses.ps1 -TenantId <dest-guid> -AdminUpn admin@contoso.onmicrosoft.com -MappingFile .\UserMapping.csv `
    -SourceSnapshot .\SourceRecipients-pre-<stamp>.csv -Domain fabrikam.com -WhatIf
.\Switch-TargetAddresses.ps1 -TenantId <dest-guid> -AdminUpn admin@contoso.onmicrosoft.com -MappingFile .\UserMapping.csv `
    -SourceSnapshot .\SourceRecipients-pre-<stamp>.csv -Domain fabrikam.com -RemoveConflictingGuests -ResetPasswords
.\Request-TargetOneDrives.ps1 -AdminUrl https://contoso-admin.sharepoint.com -AdminUpn admin@contoso.onmicrosoft.com
.\Test-TargetTenant.ps1 -TenantId <dest-guid> -AdminUpn admin@contoso.onmicrosoft.com -InventoryFolder .\Inventory-fabrikam-<date> `
    -MappingFile .\UserMapping.csv -Domain fabrikam.com -SourceSnapshot .\SourceRecipients-pre-<stamp>.csv -MailConfigFolder .\MailConfig-fabrikam.com-<date>
.\Add-TargetDomain.ps1 -TenantId <dest-guid> -AdminUpn admin@contoso.onmicrosoft.com -Domain fabrikam.com -EnableDkim
.\Get-MigrationReport.ps1 -Html
```

**Tip:** run `Get-MigrationReport.ps1` after steps 8, 12 and 16, not just at the end. A Conflict found at step 12 is a five-minute fix; found on Monday it's a week of bounced mail.
**Trap:** use a `-pre` snapshot, never a `-post` one. If `Remove-SourceDomain` finds nobody holding the domain, it names the file `-post` and warns in red, and `Switch-TargetAddresses` refuses it.

### Rollback

**Before step 8:** stop. Nothing has changed for users. Tell the business owner, send the postponement notice, and pick a new date. The copy keeps its history, so the next attempt only tops up. The destination's TXT can stay in DNS: it does nothing on its own. Remove the unverified domain from the destination if the new date is weeks away.

**After step 8:** rolling back is real work, but it is possible, and it's worth knowing you can before you start. It's realistic for the first few hours, before people have started working in the destination. After the first working morning, fixing forward is almost always cheaper. Set a time box with the business owner at the go/no-go (for example: "if mail isn't flowing in the destination by midnight, we roll back").

You need three files from tonight: the night's `SourceRecipients-pre-<stamp>.csv` (step 6), the `RemoveSourceDomain-<stamp>.csv` report from step 8 (it logs every sign-in name change), and the DNS export from step 2. And the source's old MS=ms TXT must still be in DNS (step 4 told the DNS owner to leave it).

| # | Tenant | Action |
| --- | --- | --- |
| R1 | Both | Lead and business owner agree to roll back. Note the time. Pause the copy tool's jobs |
| R2 | Destination | Release the domain. `Remove-SourceDomain.ps1` works against any tenant: point `-TenantId` at the destination. `-WhatIf` first, then `-RemoveDomain`. It moves sign-in names and addresses back to onmicrosoft.com, disables DKIM and removes the domain. The X500s it leaves behind are harmless |
| R3 | Source | Add the domain back: Microsoft 365 admin centre, Settings, Domains, Add domain, or `New-MgDomain -BodyParameter @{ Id = 'fabrikam.com' }`, then `Confirm-MgDomain -DomainId fabrikam.com`. The old MS=ms TXT is still published, so it verifies straight away. If it was removed, `Get-MgDomainVerificationDnsRecord -DomainId fabrikam.com` gives the value to publish again |
| R4 | Source | Sign-in names back, from the step 8 report. Then wait 10 to 15 minutes for Exchange to catch up (the same SyncLag the release met) |
| R5 | Source | Addresses back, from the night's `-pre` snapshot. It holds each recipient's full address list, primary included |
| R6 | DNS | Put the DKIM CNAMEs back to the source values from the step 2 export. Re-enable DKIM in the source (`Set-DkimSigningConfig -Identity fabrikam.com -Enabled $true`) once they resolve. MX, SPF and autodiscover usually never changed |
| R7 | Source | Remove any forwards or rules you changed for the night. Test mail in and out from an outside address. Test sign-in for two people with their normal passwords |
| R8 | Destination | Convert the people step 11 licensed back to shared (`Set-Mailbox -Type Shared`), or block their sign-in, so nobody starts working there. Keep the content: it's the next attempt's head start |
| R9 | Both | Comms sends the postponement notice. Anything delivered to the destination in the gap is in the destination mailboxes; copy it back or forward it by hand. Senders whose mail bounced in the gap got a bounce report and will need to resend |

```powershell
# R2 destination: release (see what it would do first)
.\Remove-SourceDomain.ps1 -TenantId <dest-guid> -AdminUpn admin@contoso.onmicrosoft.com -Domain fabrikam.com -WhatIf
.\Remove-SourceDomain.ps1 -TenantId <dest-guid> -AdminUpn admin@contoso.onmicrosoft.com -Domain fabrikam.com -RemoveDomain

# R3 to R5 source: connect as the source admin, check Get-ConnectionInformation, then
New-MgDomain -BodyParameter @{ Id = 'fabrikam.com' }
Confirm-MgDomain -DomainId fabrikam.com

# R4 sign-in names back. The report logged Area 'Sign-in', Object = old name, Detail = '-> new name'
Import-Csv .\RemoveSourceDomain-<stamp>.csv | Where-Object { $_.Area -eq 'Sign-in' -and $_.Status -eq 'Done' } | ForEach-Object {
    $now = ($_.Detail -replace '^->\s*', '').Trim()
    Update-MgUser -UserId $now -UserPrincipalName $_.Object
}
Start-Sleep -Seconds 900

# R5 addresses back. The snapshot's EmailAddresses is the full list from before the release, primary included
foreach ($r in Import-Csv .\SourceRecipients-pre-<stamp>.csv) {
    if ($r.EmailAddresses -notmatch '@fabrikam\.com') { continue }
    $list = @($r.EmailAddresses -split ';' | Where-Object { $_ })
    $set = switch -Wildcard ($r.RecipientTypeDetails) {
        'GroupMailbox' { 'Set-UnifiedGroup' }
        '*Mailbox'     { 'Set-Mailbox' }
        '*Group'       { 'Set-DistributionGroup' }
        'MailContact'  { 'Set-MailContact' }
        'MailUser'     { 'Set-MailUser' }
        default        { $null }
    }
    if (-not $set) { Write-Warning "Do $($r.DisplayName) ($($r.RecipientTypeDetails)) by hand"; continue }
    & $set -Identity $r.Alias -EmailAddresses $list -WhatIf    # remove -WhatIf once one recipient looks right
}
```

**Tip:** run R5 for one recipient first, without `-WhatIf`, and check it in the Exchange admin centre before doing the rest.
**Trap:** if R5 says an address is still the WindowsLiveId of someone else, R4 hasn't synced yet. Wait and run it again; it is safe to re-run.
**Trap:** don't delete the source's old MS=ms TXT until T+7. It costs nothing and it's the only thing that makes R3 instant.

### If something goes wrong on the night

| Problem | First thing to check |
| --- | --- |
| Destination says the domain is still in use | Removal takes a while to settle. Wait and retry `-Verify`. Check the source shows it gone |
| Remove fails with a blocker | The script lists it. Usually a deleted user in the recycle bin, a guest, or a group alias. You're still before the point where addresses are stripped if it stopped on a hard blocker |
| Mail bouncing after the release | Normal until `-Verify` succeeds. If it carries on after, check the MX the destination asked for and that the old TTL has expired |
| Switch-TargetAddresses reports Conflict | An address already used in the destination, usually by a guest. `-RemoveConflictingGuests`, or remove the object by hand and re-run |
| Switch-TargetAddresses refuses the snapshot | You gave it a `-post` file. Use the night's `-pre` file from step 6, or the one from T-1 |
| Test sign-in fails with error 53003 | Location-based Conditional Access in the destination. Test from an allowed location. The account is fine |
| A script fails with "Could not load file or assembly" | Graph module version clash. Open a fresh window. `Test-ModuleHealth.ps1 -Fix` if versions differ |
| A script says it's signed in to the wrong tenant | Re-run it with `-AdminUpn`. It disconnects and signs in as the account you name |
| Copies crawling or failing with 429 | Throttling. Let them back off; don't start more jobs. See throttling above |

## T+1: first working day

- [ ] Comms sends the cutover-morning email first thing, and the T+1 FAQ by lunchtime. If you finished well ahead of plan, the "it's done early" note instead
- [ ] Day-one support in place before people start
- [ ] Hand out temporary passwords. Help people register MFA in the destination
- [ ] Phones: help people remove and re-add their work account
- [ ] Block sign-in to the source for everyone who moved: `Block-SourceSignIn.ps1 -WhatIf`, then for real. Accounts with admin roles are reported, not blocked: do those in the admin centre, last, and keep one break-glass account
- [ ] Confirm moved users must change their temporary password: `Test-TargetTenant.ps1 ... -CutoverDate <date>` shows a Password row per person. Pass means the flag is on, or they've already changed it
- [ ] Check the `Request-TargetOneDrives` report: every person Created or Already existed. Re-run for anyone Pending
- [ ] Re-run [02b](../02b-UserMapping/) now that sign-in names have changed. Re-import the new `.sgum` into the copy tool
- [ ] Build the OneDrive copy list with [`New-DriveCopyList.ps1`](../copy-options/) (leavers into the archive library, staff into their new OneDrive), then start the OneDrive copy, outside hours if it's big
- [ ] Keep a list of every issue raised, who raised it and what fixed it
- [ ] Run 09 again at the end of the day

```powershell
# T+1
.\Block-SourceSignIn.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -MappingFile .\UserMapping.csv -WhatIf
.\Block-SourceSignIn.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -MappingFile .\UserMapping.csv
.\New-UserMapping.ps1 -SourceAdmin admin@fabrikam.onmicrosoft.com -DestinationAdmin admin@contoso.onmicrosoft.com
.\New-DriveCopyList.ps1 -SourceAdminUrl https://fabrikam-admin.sharepoint.com -Leavers -MappingFile .\UserMapping.csv `
    -DestinationSite https://contoso.sharepoint.com/sites/Archive -DestinationLibrary 'Former Staff OneDrive' -CreateFolders
```

## T+2 to T+6: settle

- [ ] Final incremental copies of mail, OneDrive, SharePoint and Teams
- [ ] Rebuild transport rules from 05's `TransportRules.xml`, disabled, then test and enable
- [ ] Repoint apps and devices that send mail (scanners, CRM, website forms)
- [ ] Work through the owner-recreates checklist in [What doesn't move](What-Doesnt-Move.md): flows, Planner, Forms, Bookings, enterprise apps
- [ ] Check DKIM is enabled and DMARC reports look normal
- [ ] Remove yourself as owner from groups 06 created, if you don't need to stay
- [ ] Delete the temporary passwords file

## T+7: sign-off

- [ ] Run 09 one last time. Attach the CSV to the handover
- [ ] Walk the business owner through the [handover](Handover-Template.md)
- [ ] Remove the source's old MS=ms TXT from DNS, now rollback is off the table
- [ ] Agree the decommission date for the source, and what happens to its data before then ([10](../10-Decommission/)). Check it against the source renewal date in the licensing timeline
- [ ] Close or hand over every open item on the issue list
- [ ] Business owner signs off

## Three ways to run cutover night

| Route | How |
| --- | --- |
| Paid tool | The copy tool handles the final incrementals and its own mapping. The domain move, addresses, X500s and DNS are still yours: most tools don't do those for you between two Microsoft 365 tenants |
| Microsoft's own | Admin centres: remove the domain in the source (Settings, Domains), add and verify it in the destination, then edit addresses user by user in Exchange admin. Works for ten people. X500s need PowerShell either way |
| By hand, scripted | The toolkit's 07 scripts in the order above. Each one checks the tenant, supports `-WhatIf` and writes a report |

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 1.2.0 | 2026-11-10 | Made this the canonical order. Snapshot at T-1 and again on the night before readiness; `-BlockersOnly` for blockers; DNS full scan and any NS move at T-14; `-FixApplications`, `-RemoveConflictingGuests`, `-SharedMailboxesOnly`, 02b and 04 placed in the order. Snapshot name is `SourceRecipients-pre-<stamp>.csv`. Added licensing timeline, coexistence, throttling and a worked rollback after the release |
| 1.1.0 | 2026-11-10 | Add-first T-0 order. T-14 DNS host check and name server move. T-1 readiness check and test sign-in plan. T+1 block source sign-in, password-change check, rebuild the mapping. From a live cutover |
| 1.0.0 | 2026-11-09 | First release |
