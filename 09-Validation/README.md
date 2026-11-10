# Test-TargetTenant

**TL;DR:** A read-only check of the destination against everything the earlier parts recorded. For every row in your mapping file it checks the recipient exists, has the right type, licence and archive, has about as many items as the source did, and after cutover has its old address, sign-in name and X500 address. It checks people must change their temporary password, and flags Conditional Access that will block your own test sign-ins and a missing MFA policy. It can also check OneDrives and distribution list membership. You get a CSV of Pass, Warn and Fail rows and a summary table. Nothing in the tenant is changed.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.1.0 |
| Released | 10 November 2026 |
| Licence | MIT |

Part 9 of my [M365 tenant-to-tenant migration toolkit](../README.md). Run it after the main copy, again on cutover night once [07](../07-DomainCutover/) has finished, and once more a few days later.

## Why you need this

By cutover night you've run a dozen scripts and a copy tool across two tenants. Each one wrote a report, and each report had a few skips or problems in it. It's very easy to miss one.

This pulls it together. It reads the inventory from [01](../01-SourceInventory/), the plan from [02](../02-TargetMailboxes/), your mapping file, the snapshot from [07](../07-DomainCutover/) and the export from [05](../05-MailConfig/), and checks the destination against all of them in one go. The result is one list of what's right and what isn't, which is also what I hand to the business owner when they ask "is it done?"

## What happens without it

- **Someone is still staged as shared.** They can't sign in on Monday, and you find out from them.
- **An archive never got enabled.** The copy tool skipped the archive content quietly, and the old tenant is due to be switched off.
- **A licence went on the wrong SKU,** or not at all, and features people relied on aren't there.
- **The X500 step missed a few people.** Their colleagues' replies to old mail bounce for weeks, and it looks random.
- **Half a mailbox came across.** A copy job stopped part way and the summary screen still looked green.
- **A distribution list is missing members,** and the team address quietly stops reaching some of the team.

## What it checks

| Check | What it compares | Needs |
| --- | --- | --- |
| Exists | The destination in the mapping file resolves to a recipient | Always |
| Type | `UserMailbox` or `SharedMailbox` per the plan's Kind, otherwise the source type | Always. Plan for the strict version |
| Licence | The plan's SKU is assigned | `TargetBuildPlan.csv` |
| Archive | Archive is active where the source had one | Always |
| Items, ArchiveItems | Item counts within `-Tolerance` percent (default 5) of `04-Mailboxes.csv` | Always |
| Address | Plan rows: old address is primary. Everyone else: old address present | `-Domain` |
| SignIn | Plan rows with a mailbox sign in with their old address | `-Domain` and the plan |
| X500 | The source LegacyExchangeDN is present as an X500 address | `-SourceSnapshot` |
| OneDrive | People (Kind User) have a OneDrive for their sign-in name | `-SharePointAdminUrl` and the plan |
| Mapping | Plan rows that aren't in the mapping file | The plan |
| ListMembers | Distribution list member counts against the 05 export | `-MailConfigFolder` |
| Password | People (Kind User) have "must change password at next sign-in" on. With `-CutoverDate`, a password already changed after cutover passes as informational | The plan |
| CA-Location | Every enabled or report-only Conditional Access policy with a location condition, as a Warn | Policy.Read.All. Skip with `-SkipTenantChecks` |
| MFA | Warn if no enabled policy requires MFA for all users and security defaults are off. A finding, nothing is changed | Policy.Read.All |

How the results are judged:

- **Pass** means it matches.
- **Warn** means it's different and might be fine, for example more items than the source (new mail since the inventory) or a list with fewer members because some had no mapping. Read each one.
- **Fail** means something needs doing. The Detail column says what, and usually which part to re-run.

## Requirements

- `ExchangeOnlineManagement` 3.x, `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users`, `Microsoft.Graph.Identity.DirectoryManagement`
- Global Reader in the destination is enough for everything except OneDrives
- For the OneDrive check: `Microsoft.Online.SharePoint.PowerShell` and SharePoint Administrator
- The files from earlier parts. Only the inventory folder and the mapping file are required

## Usage

```powershell
# After the main copy, before cutover
.\Test-TargetTenant.ps1 -TenantId <destination-tenant-guid> -InventoryFolder .\Inventory-fabrikam-<date> `
    -MappingFile .\UserMapping.csv -MailConfigFolder .\MailConfig-fabrikam.com-<date>

# After cutover: everything
.\Test-TargetTenant.ps1 -TenantId <destination-tenant-guid> -InventoryFolder .\Inventory-fabrikam-<date> `
    -MappingFile .\UserMapping.csv -Domain fabrikam.com, fabrikam.co.uk `
    -SourceSnapshot .\SourceRecipients-<date>.csv -MailConfigFolder .\MailConfig-fabrikam.com-<date> `
    -SharePointAdminUrl https://contoso-admin.sharepoint.com

# Stricter on item counts
.\Test-TargetTenant.ps1 ... -Tolerance 2

# The morning after: name the account, and tell it when cutover was
.\Test-TargetTenant.ps1 -TenantId <destination-tenant-guid> -AdminUpn admin@contoso.onmicrosoft.com `
    -InventoryFolder .\Inventory-fabrikam-<date> -MappingFile .\UserMapping.csv -Domain fabrikam.com -CutoverDate 2026-11-10
```

**Tip: name the account.** With both tenants open on one PC, `Connect-ExchangeOnline` and `Connect-MgGraph` happily reuse whichever session is already there. Add `-AdminUpn admin@contoso.onmicrosoft.com` and the script disconnects any session for another account or tenant, signs in to Exchange Online and Graph as the one you named, and prints `Signed in as X to tenant Y`. On the live cutover this would have saved a confusing error at the worst moment.

Run it from the folder with `TargetBuildPlan.csv`, or pass `-PlanFile`.

## Output

- **On screen:** every Warn and Fail as it's found, then a table of Pass, Warn and Fail per check.
- **CSV:** `Validation-<domain>-<date>.csv`, one row per check per recipient, with Object, Check, Result, Expected, Actual and Detail.

The CSV contains names and addresses. Keep it out of source control and shared drives. The repo's `.gitignore` already excludes `*.csv`.

## Lessons learnt

- **Re-run the inventory before you validate.** Item counts are compared with `04-Mailboxes.csv`, which is a snapshot from the day you ran part 1. If that was a month ago, the source has grown and everything looks short. Run part 1 again close to cutover and point this at the new folder.
- **Item counts are a guide, not proof.** Copy tools don't always move every item type, and they count things differently. That's why there's a tolerance. If a mailbox is well outside it, open the copy tool's report for that mailbox before re-copying anything.
- **More items isn't always good news.** After cutover, new mail pushes the count up and that's fine. Before cutover, a mailbox well above the source can mean a re-copy made duplicates. Worth a look.
- **Keep the mapping file honest.** Everything is checked through it. If someone was added to the plan late and never made it into the mapping, the Mapping check flags them, but nothing else about them is checked until you add the row.
- **Run it more than once.** Before cutover it tells you what's left to copy. On the night it tells you whether cutover worked. A few days later it catches anything that drifted, like a licence someone removed.
- **The password flag matters.** A temporary password handed out on cutover night should stop working the first time it's used. If the change-at-next-sign-in flag is off, it stays valid until someone changes it, and it was probably written on paper. Graph returns an empty value rather than False when the flag is off, so the script treats empty as off.
- **Test sign-ins from where the users are.** The destination on the live cutover had a "block outside allowed countries" policy. I was testing from abroad, and every test sign-in failed with error 53003. The accounts were fine. The CA-Location warning is there so you know before you panic.
- **Report the MFA gap, don't fix it.** No MFA policy is a real finding, and it goes in the handover for the business owner. Changing security policy in a tenant that hosts other companies isn't part of a migration.
- **Use it for sign-off.** The summary table is the closest thing to a straight answer when the business asks if it's finished. I attach the CSV to the handover.

## Three ways to do it

| Route | How |
| --- | --- |
| Paid tool | ShareGate's task reports show what it copied and what it skipped, item by item. They say nothing about addresses, X500s, licences or sign-in. Use both |
| Microsoft's own | Exchange admin centre and Entra admin centre, person by person: type, licence, addresses, sign-in name. `Get-EXOMailboxStatistics` for counts. Sign-in logs for 53003 failures |
| By hand | A spreadsheet of every person with a column per check, filled in from the admin centres. Accurate for ten people, if you have an afternoon |

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 1.1.0 | 2026-11-10 | `-AdminUpn`. Password, CA-Location and MFA checks. `-CutoverDate`, `-SkipTenantChecks` |
| 1.0.0 | 2026-11-09 | First release |
