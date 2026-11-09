# Cutover runbook

**TL;DR:** The two weeks either side of cutover, day by day, with the exact toolkit step for each. It has go/no-go checks before the point of no return, a clear rollback point, and who does what. Copy it, put your dates and names in, and agree it with the business owner before T-14. Examples use fabrikam.com moving into the contoso tenant.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.0.0 |
| Released | 9 November 2026 |
| Licence | MIT |

Part of the [docs](README.md) for my [M365 tenant-to-tenant migration toolkit](../README.md).

## Who does what

| Role | Who | Does |
| --- | --- | --- |
| Migration lead | [name] | Runs the scripts and copy jobs. Owns this runbook |
| Business owner | [name] | Makes the go/no-go call. Signs off at T+7 |
| DNS owner | [name, phone] | Changes DNS records on the night. Must be reachable during the cutover window |
| Comms owner | [name] | Sends the staff emails in [Comms templates](Comms-Templates.md) |
| Day-one support | [names] | On hand on the first working morning for sign-ins, phones and Outlook |
| Device support | [names] | Rejoins or resets laptops and phones, if that's in scope |

Swap names in. One person can hold several roles in a small business, but the DNS owner and the business owner must be named and reachable on the night.

## Before T-14

These should already be done. If they aren't, the dates move.

- [ ] Discovery done and agreed ([Discovery questionnaire](Discovery-Questionnaire.md))
- [ ] Copy option chosen ([copy-options](../copy-options/))
- [ ] Destination licences bought and visible in the destination
- [ ] Mapping file built (`Source,Destination`) and checked line by line
- [ ] [Risk register](Risk-Register.md) reviewed with the business owner

## T-14: build and first copy

| Step | Who | Action |
| --- | --- | --- |
| 1 | Lead | Run [01](../01-SourceInventory/) against the source. Keep the folder |
| 2 | Lead | Run [02](../02-TargetMailboxes/) `-Plan`, review `TargetBuildPlan.csv` with the business owner, then `-Apply -StageOnly` against the destination |
| 3 | Lead | If your tool copies archives over EWS, run [03](../03-EwsArchiveReadiness/) with `-Fix` in both tenants |
| 4 | Lead | Run [06](../06-GroupsTeamsSites/) `-Export` then `-Import` so groups and Teams exist before you copy into them |
| 5 | Lead | Import the user mapping into the copy tool. Start the first mail and archive copy |
| 6 | Lead | Run [07](../07-DomainCutover/) `Export-DomainDns.ps1` and send the DNS owner a list of TTLs to lower |
| 7 | Comms | Agree the cutover date in writing with the business owner |

```powershell
.\Invoke-SourceTenantInventory.ps1 -SharePointAdminUrl https://fabrikam-admin.sharepoint.com
.\New-TargetMailboxes.ps1 -Plan -TargetDomain contoso.onmicrosoft.com
.\New-TargetMailboxes.ps1 -Apply -TargetDomain contoso.onmicrosoft.com -UsageLocation GB -StageOnly
.\Export-DomainDns.ps1 -Domain fabrikam.com -Server 1.1.1.1
```

## T-7: config and heads-up

| Step | Who | Action |
| --- | --- | --- |
| 1 | Lead | Check the first copy finished. Fix and re-run anything that failed |
| 2 | Lead | Run [05](../05-MailConfig/) `-Export` against the source, then `-Import -WhatIf` and `-Import` against the destination |
| 3 | Lead | Start SharePoint and Teams copies into the groups 06 created |
| 4 | Lead | Run [07](../07-DomainCutover/) `Remove-SourceDomain.ps1 -SnapshotOnly`, then `-WhatIf`, and clear every blocker it lists (deleted users, leftover aliases) |
| 5 | Lead | Run [09](../09-Validation/) without `-Domain`. Expect warnings; read every fail |
| 6 | Comms | Send the T-7 heads-up email |
| 7 | DNS owner | Lower TTLs on MX, autodiscover, SPF and any CNAMEs flagged |

## T-3: top up and dry run

- [ ] Incremental copy of mail and archives
- [ ] Re-run 01 so item counts for validation are recent
- [ ] Re-run 09 and compare with T-7. Fails should be going down
- [ ] Walk through the cutover night steps with the DNS owner, out loud, with times
- [ ] Confirm day-one support people and where they'll be
- [ ] Prepare the temporary password hand-out method (in person, by phone, or by manager). Not by email to the old mailbox

## T-1: go/no-go

The business owner makes the call. Don't start cutover with a no.

| Check | Go if |
| --- | --- |
| Copy status | Every mailbox has had at least one complete pass. Archives done, or agreed to finish after cutover |
| Validation | 09 has no unexplained fails. Every remaining fail has an owner and a plan |
| Licences | Enough seats for everyone in the plan |
| Blockers | `Remove-SourceDomain -WhatIf` shows no blockers left |
| DNS | TTLs were lowered at least a day ago. DNS owner confirmed for tonight or tomorrow |
| People | Day-one support confirmed. Comms owner ready to send the T-1 email |
| Business | No month end, payroll run or key event tomorrow |

If it's a go:

- [ ] Comms sends the T-1 email
- [ ] Lead records the go decision, time and who made it

## T-0: cutover night

Do these in order. Note the time against each step.

| # | Time | Who | Action |
| --- | --- | --- | --- |
| 1 | | Lead | Final incremental copy of mail. Let it finish |
| 2 | | Lead | `Export-DomainDns.ps1` again. This file is your rollback copy of DNS |
| 3 | | Lead | `Remove-SourceDomain.ps1 -SnapshotOnly`. Keep the CSV safe |
| | | **Lead and business owner** | **Rollback point. Last chance to stop. Before this line, nothing has changed for users** |
| 4 | | Lead | `Remove-SourceDomain.ps1 -RemoveDomain`. The domain leaves the source. Mail to it can bounce or queue until step 7 |
| 5 | | Lead | Disconnect from the source. `Add-TargetDomain.ps1` in the destination. Send the TXT record to the DNS owner |
| 6 | | DNS owner | Add the TXT record |
| 7 | | Lead | `Add-TargetDomain.ps1 -Verify`. Send the MX, SPF, autodiscover and DKIM records it prints to the DNS owner |
| 8 | | DNS owner | Change MX, SPF, autodiscover and add the DKIM CNAMEs |
| 9 | | Lead | `New-TargetMailboxes.ps1 -Apply` (no `-StageOnly`): people become user mailboxes and licensed. Run twice, about 15 minutes apart |
| 10 | | Lead | `Switch-TargetAddresses.ps1 -WhatIf`, read it, then run with `-ResetPasswords`. These passwords replace any step 9 set, so hand out the ones from this file, with the new sign-in names |
| 11 | | Lead | `Request-TargetOneDrives.ps1` ([04](../04-TargetOneDrives/)) for anyone still without one |
| 12 | | Lead | Send a test message in and out from an outside address. Check it lands |
| 13 | | Lead | `Test-TargetTenant.ps1` ([09](../09-Validation/)) with `-Domain` and `-SourceSnapshot` |
| 14 | | Lead | `Add-TargetDomain.ps1 -EnableDkim` once the CNAMEs resolve. Can wait until morning |
| 15 | | Lead | Final incremental copy of mail, now that addresses are final |
| 16 | | Lead | Tell the business owner it's done, with the 09 summary |

```powershell
.\Export-DomainDns.ps1 -Domain fabrikam.com -Server 1.1.1.1
.\Remove-SourceDomain.ps1 -TenantId <source-guid> -Domain fabrikam.com -SnapshotOnly
# ---- rollback point ----
.\Remove-SourceDomain.ps1 -TenantId <source-guid> -Domain fabrikam.com -RemoveDomain
Disconnect-ExchangeOnline -Confirm:$false; Disconnect-MgGraph
.\Add-TargetDomain.ps1 -TenantId <dest-guid> -Domain fabrikam.com
.\Add-TargetDomain.ps1 -TenantId <dest-guid> -Domain fabrikam.com -Verify
.\New-TargetMailboxes.ps1 -Apply -TargetDomain contoso.onmicrosoft.com -UsageLocation GB
.\Switch-TargetAddresses.ps1 -TenantId <dest-guid> -MappingFile .\UserMapping.csv -SourceSnapshot .\SourceRecipients-<date>.csv -Domain fabrikam.com -ResetPasswords
.\Test-TargetTenant.ps1 -TenantId <dest-guid> -InventoryFolder .\Inventory-fabrikam-<date> -MappingFile .\UserMapping.csv `
    -Domain fabrikam.com -SourceSnapshot .\SourceRecipients-<date>.csv -MailConfigFolder .\MailConfig-fabrikam.com-<date>
```

### Rollback

**Before step 4:** stop. Nothing has changed for users. Tell the business owner, send a short "we've postponed" note, and pick a new date. The copy keeps its history, so the next attempt only tops up.

**After step 4:** there's no clean rollback. The source has let go of the domain, and getting it back means reversing every step in the source by hand. Treat step 4 as the point of no return. If the destination step fails, keep going forward and fix it: the domain can only be verified in one place, and the destination is where it's heading. Keep the DNS export from step 2 in case records need restoring for anything else that used them.

### If something goes wrong on the night

| Problem | First thing to check |
| --- | --- |
| Destination says the domain is still in use | Removal takes a while to settle. Wait and retry `-Verify`. Check the source shows it gone |
| Remove fails with a blocker | The script lists it. Usually a deleted user in the recycle bin, a guest, or a group alias |
| Mail bouncing after MX change | Check the MX record the destination asked for, and that the old TTL has expired |
| Switch-TargetAddresses reports problems | Read the report. Usually a recipient missing from the mapping, or an address already used elsewhere |

## T+1: first working day

- [ ] Comms sends the cutover-morning email first thing, and the T+1 FAQ by lunchtime
- [ ] Day-one support in place before people start
- [ ] Hand out temporary passwords. Help people register MFA in the destination
- [ ] Phones: help people remove and re-add their work account
- [ ] Start the OneDrive copy now that sign-in names are final
- [ ] Keep a list of every issue raised, who raised it and what fixed it
- [ ] Run 09 again at the end of the day

## T+2 to T+6: settle

- [ ] Final incremental copies of mail, OneDrive, SharePoint and Teams
- [ ] Rebuild transport rules from 05's `TransportRules.xml`, disabled, then test and enable
- [ ] Repoint apps and devices that send mail (scanners, CRM, website forms)
- [ ] Check DKIM is enabled and DMARC reports look normal
- [ ] Remove yourself as owner from groups 06 created, if you don't need to stay
- [ ] Delete the temporary passwords file

## T+7: sign-off

- [ ] Run 09 one last time. Attach the CSV to the handover
- [ ] Walk the business owner through the [handover](Handover-Template.md)
- [ ] Agree the decommission date for the source, and what happens to its data before then
- [ ] Close or hand over every open item on the issue list
- [ ] Business owner signs off
