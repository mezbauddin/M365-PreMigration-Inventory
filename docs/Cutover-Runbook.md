# Cutover runbook

**TL;DR:** The two weeks either side of cutover, day by day, with the exact toolkit step for each. It has go/no-go checks before the point of no return, a clear rollback point, and who does what. Copy it, put your dates and names in, and agree it with the business owner before T-14. Examples use fabrikam.com moving into the contoso tenant.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.1.0 |
| Released | 10 November 2026 |
| Licence | MIT |

Part of the [docs](README.md) for my [M365 tenant-to-tenant migration toolkit](../README.md).

## Who does what

| Role | Who | Does |
| --- | --- | --- |
| Migration lead | [name] | Runs the scripts and copy jobs. Owns this runbook |
| Business owner | [name] | Makes the go/no-go call. Signs off at T+7 |
| DNS owner | [name, phone] | Can log in to the **DNS host** (where the name servers point), not just the registrar. Changes DNS records on the night. Must be reachable during the cutover window |
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
| 7 | Lead and DNS owner | Confirm who can log in to the **DNS host**. Check with `Resolve-DnsName fabrikam.com -Type NS`: if the name servers aren't at the company whose login you have, move DNS hosting now, at least 48 hours before cutover ([DNS move playbook](DNS-Move-Playbook.md)) |
| 8 | Lead | Run [tools/Test-ModuleHealth.ps1](../tools/) on the PC you'll use on the night. Fix any Graph version mismatch now, not on the night |
| 9 | Comms | Agree the cutover date in writing with the business owner |

**Trap:** the registrar and the DNS host are often different companies. Having the registrar login proves nothing about who can add a TXT record on the night.

```powershell
.\Invoke-SourceTenantInventory.ps1 -SharePointAdminUrl https://fabrikam-admin.sharepoint.com
.\New-TargetMailboxes.ps1 -Plan -TargetDomain contoso.onmicrosoft.com
.\New-TargetMailboxes.ps1 -Apply -TargetDomain contoso.onmicrosoft.com -UsageLocation GB -StageOnly -AdminUpn admin@contoso.onmicrosoft.com
.\Export-DomainDns.ps1 -Domain fabrikam.com -Server 1.1.1.1
Resolve-DnsName fabrikam.com -Type NS      # who really hosts DNS?
.\Test-ModuleHealth.ps1
# Only if the DNS host is out of reach: full scan, recreate at the registrar, switch NS (see the DNS move playbook)
.\Export-DomainDns.ps1 -Domain fabrikam.com -FullScan
.\Set-DomainDnsRecords.ps1 -Domain fabrikam.com -Provider GoDaddy -FromScan .\DomainDns-FullScan-<date>.csv -WhatIf
```

## T-7: config and heads-up

| Step | Who | Action |
| --- | --- | --- |
| 1 | Lead | Check the first copy finished. Fix and re-run anything that failed |
| 2 | Lead | Run [05](../05-MailConfig/) `-Export` against the source, then `-Import -WhatIf` and `-Import` against the destination |
| 3 | Lead | Start SharePoint and Teams copies into the groups 06 created |
| 4 | Lead | Run [07](../07-DomainCutover/) `Remove-SourceDomain.ps1 -SnapshotOnly`, then `-WhatIf`, and clear every blocker it lists (deleted users, proxy-only users, leftover aliases, apps with URLs on the domain). `Clear-DeletedUserDomain.ps1` handles deleted users |
| 4a | Lead | Look for guests and other objects in the destination already holding source addresses. They block `Switch-TargetAddresses` at cutover. `Test-CutoverReadiness.ps1` lists them |
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

Run the readiness check first. It runs the module check, reads both tenants and DNS, and gives PASS, WARN and FAIL for each blocker on this page:

```powershell
.\Test-CutoverReadiness.ps1 -SourceTenantId <source-guid> -DestinationTenantId <dest-guid> `
    -SourceAdmin admin@fabrikam.onmicrosoft.com -DestinationAdmin admin@contoso.onmicrosoft.com `
    -Domain fabrikam.com -MappingFile .\UserMapping.csv -SourceSnapshot .\SourceRecipients-<date>.csv
.\Get-MigrationReport.ps1 -Html      # every report so far, on one page
```

| Check | Go if |
| --- | --- |
| Readiness | `Test-CutoverReadiness.ps1` has no FAIL. Every WARN read and accepted |
| Copy status | Every mailbox has had at least one complete pass. Archives done, or agreed to finish after cutover |
| Validation | 09 has no unexplained fails. Every remaining fail has an owner and a plan |
| Licences | Enough seats for everyone in the plan |
| Blockers | `Remove-SourceDomain -WhatIf` shows no blockers left |
| DNS | TTLs were lowered at least a day ago. DNS owner can log in to the DNS host (tested, not assumed) and is confirmed for the night. Any name server move finished at least 48 hours ago |
| Test sign-in plan | Who tests the moved accounts, from where. If the destination has location-based Conditional Access, the tester must be in an allowed location, or their test account excluded for the night and the exclusion removed after. Otherwise sign-in fails with error 53003 and looks like a broken migration |
| People | Day-one support confirmed. Comms owner ready to send the T-1 email |
| Business | No month end, payroll run or key event tomorrow |

If it's a go:

- [ ] Comms sends the T-1 email
- [ ] Lead records the go decision, time and who made it

## T-0: cutover night

Do these in order. Note the time against each step. This is the **add-first** order: the destination gives you its verification TXT while the source still holds the domain, so the TXT is published and resolving before anything changes for users. Only verification has to wait for the source to let go.

| # | Time | Who | Action |
| --- | --- | --- | --- |
| 1 | | Lead | Final incremental copy of mail. Let it finish |
| 2 | | Lead | `Export-DomainDns.ps1` again. This file is your rollback copy of DNS |
| 3 | | Lead | Destination: `Add-TargetDomain.ps1`. Send the TXT record it prints to the DNS owner |
| 4 | | DNS owner | Add the TXT record. Leave everything else alone |
| 5 | | Lead | `Add-TargetDomain.ps1 -WaitForDns` until the TXT resolves |
| 6 | | Lead | Source: `Remove-SourceDomain.ps1 -SnapshotOnly`. Keep the CSV safe. This is the snapshot that counts |
| | | **Lead and business owner** | **Rollback point. Last chance to stop. Before this line, nothing has changed for users. The TXT record does no harm if you stop** |
| 7 | | Lead | `Remove-SourceDomain.ps1 -RemoveDomain`. The domain leaves the source. Mail to it can bounce or queue until step 8 |
| 8 | | Lead | Disconnect from the source. `Add-TargetDomain.ps1 -Verify`. If it says the domain is still in use, wait a few minutes and retry. Send the records it marks Change or Add to the DNS owner. Records marked Unchanged stay as they are, and never apply anything marked DoNotApply |
| 9 | | DNS owner | Change only what step 8 asked for: often just autodiscover and the DKIM CNAMEs. MX and SPF are frequently the same between two Microsoft 365 tenants |
| 10 | | Lead | `New-TargetMailboxes.ps1 -Apply` (no `-StageOnly`): people become user mailboxes and licensed. Run twice, about 15 minutes apart |
| 11 | | Lead | `Switch-TargetAddresses.ps1 -WhatIf`, read it, then run with `-ResetPasswords`. These passwords replace any step 10 set, so hand out the ones from this file, with the new sign-in names |
| 12 | | Lead | `Request-TargetOneDrives.ps1` ([04](../04-TargetOneDrives/)) for anyone still without one |
| 13 | | Lead | Send a test message in and out from an outside address. Check it lands |
| 14 | | Tester | Test sign-in for two moved users, from an allowed location (see the T-1 test sign-in plan) |
| 15 | | Lead | `Test-TargetTenant.ps1` ([09](../09-Validation/)) with `-Domain` and `-SourceSnapshot` |
| 16 | | Lead | `Add-TargetDomain.ps1 -EnableDkim` once the CNAMEs resolve. Can wait until morning |
| 17 | | Lead | Final incremental copy of mail, now that addresses are final |
| 18 | | Lead | `Get-MigrationReport.ps1 -Html`. Tell the business owner it's done, with the 09 summary and the report |

```powershell
# Every tenant-changing script takes -AdminUpn: it disconnects any other session and prints who and where before changing anything
.\Export-DomainDns.ps1 -Domain fabrikam.com -Server 1.1.1.1
.\Add-TargetDomain.ps1 -TenantId <dest-guid> -AdminUpn admin@contoso.onmicrosoft.com -Domain fabrikam.com
.\Add-TargetDomain.ps1 -TenantId <dest-guid> -AdminUpn admin@contoso.onmicrosoft.com -Domain fabrikam.com -WaitForDns
.\Remove-SourceDomain.ps1 -TenantId <source-guid> -Domain fabrikam.com -SnapshotOnly
# ---- rollback point ----
.\Remove-SourceDomain.ps1 -TenantId <source-guid> -Domain fabrikam.com -RemoveDomain
Disconnect-ExchangeOnline -Confirm:$false; Disconnect-MgGraph
.\Add-TargetDomain.ps1 -TenantId <dest-guid> -AdminUpn admin@contoso.onmicrosoft.com -Domain fabrikam.com -WaitForDns -Verify
.\New-TargetMailboxes.ps1 -Apply -TargetDomain contoso.onmicrosoft.com -UsageLocation GB -AdminUpn admin@contoso.onmicrosoft.com
.\Switch-TargetAddresses.ps1 -TenantId <dest-guid> -MappingFile .\UserMapping.csv -SourceSnapshot .\SourceRecipients-<date>.csv -Domain fabrikam.com -ResetPasswords
.\Test-TargetTenant.ps1 -TenantId <dest-guid> -AdminUpn admin@contoso.onmicrosoft.com -InventoryFolder .\Inventory-fabrikam-<date> `
    -MappingFile .\UserMapping.csv -Domain fabrikam.com -SourceSnapshot .\SourceRecipients-<date>.csv -MailConfigFolder .\MailConfig-fabrikam.com-<date>
.\Get-MigrationReport.ps1 -Html
```

**Tip:** run `Get-MigrationReport.ps1` after steps 7, 11 and 15, not just at the end. A Conflict found at step 11 is a five-minute fix; found on Monday it's a week of bounced mail.

### Rollback

**Before step 7:** stop. Nothing has changed for users. Tell the business owner, send a short "we've postponed" note, and pick a new date. The copy keeps its history, so the next attempt only tops up.

**After step 7:** there's no clean rollback. The source has let go of the domain, and getting it back means reversing every step in the source by hand. Treat step 7 as the point of no return. If the destination step fails, keep going forward and fix it: the domain can only be verified in one place, and the destination is where it's heading. Keep the DNS export from step 2 in case records need restoring for anything else that used them. The verification TXT from step 4 can stay in DNS if you stop: it does nothing on its own.

### If something goes wrong on the night

| Problem | First thing to check |
| --- | --- |
| Destination says the domain is still in use | Removal takes a while to settle. Wait and retry `-Verify`. Check the source shows it gone |
| Remove fails with a blocker | The script lists it. Usually a deleted user in the recycle bin, a guest, or a group alias |
| Mail bouncing after MX change | Check the MX record the destination asked for, and that the old TTL has expired |
| Switch-TargetAddresses reports problems | Read the report. Usually a recipient missing from the mapping, or an address already used elsewhere, often by a guest |
| Test sign-in fails with error 53003 | Location-based Conditional Access in the destination. Test from an allowed location. The account is fine |
| A script fails with "Could not load file or assembly" | Graph module version clash. Open a fresh window. `Test-ModuleHealth.ps1 -Fix` if versions differ |
| A script says it's signed in to the wrong tenant | Re-run it with `-AdminUpn`. It disconnects and signs in as the account you name |

## T+1: first working day

- [ ] Comms sends the cutover-morning email first thing, and the T+1 FAQ by lunchtime
- [ ] Day-one support in place before people start
- [ ] Hand out temporary passwords. Help people register MFA in the destination
- [ ] Phones: help people remove and re-add their work account
- [ ] Block sign-in to the source for everyone who moved: `Block-SourceSignIn.ps1 -WhatIf`, then for real. Accounts with admin roles are reported, not blocked: do those in the admin centre, last, and keep one break-glass account
- [ ] Confirm moved users must change their temporary password: `Test-TargetTenant.ps1 ... -CutoverDate <date>` shows a Password row per person. Pass means the flag is on, or they've already changed it
- [ ] Re-run [02b](../02b-UserMapping/) now that sign-in names have changed. Re-import the new `.sgum` into the copy tool
- [ ] Build the OneDrive copy list with [`New-DriveCopyList.ps1`](../copy-options/) (leavers into the archive library, staff into their new OneDrive), then start the OneDrive copy
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
- [ ] Check DKIM is enabled and DMARC reports look normal
- [ ] Remove yourself as owner from groups 06 created, if you don't need to stay
- [ ] Delete the temporary passwords file

## T+7: sign-off

- [ ] Run 09 one last time. Attach the CSV to the handover
- [ ] Walk the business owner through the [handover](Handover-Template.md)
- [ ] Agree the decommission date for the source, and what happens to its data before then
- [ ] Close or hand over every open item on the issue list
- [ ] Business owner signs off

## Three ways to run cutover night

| Route | How |
| --- | --- |
| Paid tool | The copy tool handles the final incrementals and its own mapping. The domain move, addresses, X500s and DNS are still yours: no tool does those for you between two Microsoft 365 tenants |
| Microsoft's own | Admin centres: remove the domain in the source (Settings, Domains), add and verify it in the destination, then edit addresses user by user in Exchange admin. Works for ten people. X500s need PowerShell either way |
| By hand, scripted | The toolkit's 07 scripts in the order above. Each one checks the tenant, supports `-WhatIf` and writes a report |

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 1.1.0 | 2026-11-10 | Add-first T-0 order. T-14 DNS host check and name server move. T-1 readiness check and test sign-in plan. T+1 block source sign-in, password-change check, rebuild the mapping. From a live cutover |
| 1.0.0 | 2026-11-09 | First release |
