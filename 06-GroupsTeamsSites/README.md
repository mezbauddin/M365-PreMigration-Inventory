# Copy-CollaborationMap

**TL;DR:** Before you copy Teams and SharePoint, you want the groups to already exist in the destination, on the right domain, with the site URL you chose and the right people in them. This script exports every Microsoft 365 group, Team and SharePoint site from the source, then creates the missing groups in the destination, turns them into Teams if you ask, and adds owners and members through your mapping file. Your copy tool then copies into groups that are already set up. Everything supports `-WhatIf`.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.1.0 |
| Released | 9 November 2026 |
| Licence | MIT |

Part 6 of my [M365 tenant-to-tenant migration toolkit](../README.md). Run the export any time. Run the import once the destination mailboxes exist ([02](../02-TargetMailboxes/)) and before you start copying Teams or SharePoint.

## Why you need this

Most copy tools will happily create a group or a Team for you as part of the copy. The trouble is you get whatever address domain and site URL the tool chooses. If a team with that name already exists, you can end up with a second one and a site URL ending in 2.

It's much calmer to create the groups yourself first. You pick the alias, the address domain and the visibility. Owners and members are in place before any content arrives. Then you tell the copy tool to copy into the existing group.

The export is useful on its own too. It's the list you'll want when you're scoping the job: how many Teams, how big each site is, who owns what, and which sites aren't attached to a group at all.

## What happens without it

- **Groups land on the wrong domain.** A group created by the copy tool can end up with an onmicrosoft.com address, and nobody notices until cutover, when the old address has nowhere sensible to go.
- **Site URLs you didn't choose.** Once a group's site exists, changing its URL is a job in itself. Better to get it right at creation.
- **Empty Teams.** If the user mapping isn't in place, content arrives and the members don't. People open Teams on day one and can't see their team.
- **Duplicate teams.** Copying 'as a new team' when one already exists can give you two teams with the same name.
- **Sites get forgotten.** Communication sites and old team sites that aren't attached to a group don't show up in a list of Teams. If nobody lists them, nobody copies them.

## What it does

| Part | Export | Import |
| --- | --- | --- |
| Microsoft 365 groups | Name, alias, primary address, all smtp addresses, visibility, Teams-enabled, site URL, storage, description | Created with `New-UnifiedGroup` on `-AddressDomain`, owner set to you, welcome messages off. Tagged for cutover |
| Teams | Flag on each group | With `-EnableTeams`, enabled through Graph (`PUT /groups/{id}/team`) where the source was a Team |
| Owners and members | One row per link, by primary address | Mapped through your CSV and added with `Add-UnifiedGroupLinks`, members before owners |
| Other SharePoint sites | URL, title, template, owner, storage, last modified (needs `-SharePointAdminUrl`) | Not created. Use the list to plan them |
| Private and shared channels | Their sites are listed, marked `PrivateOrSharedChannel` | Not created |

Every group the import creates is tagged with `CustomAttribute15 = T2T:<source address>`. Re-runs use that to skip what's already done, and [Switch-TargetAddresses](../07-DomainCutover/) uses it at cutover to give each group its old address.

If a group already exists in the destination and you want to use it, add a row for it to the mapping file: source group address to destination group address. The script then adds the members to that group instead of creating a new one. Don't add rows for groups you want the script to create. The tag looks after those.

## Requirements

- `ExchangeOnlineManagement` 3.x, as Exchange Administrator in each tenant
- For storage and the site list: `Microsoft.Online.SharePoint.PowerShell` and SharePoint Administrator in the source
- For `-EnableTeams`: `Microsoft.Graph.Authentication`, consenting to `Group.ReadWrite.All` in the destination
- A mapping CSV with `Source` and `Destination` columns. The same one [05](../05-MailConfig/) and [07](../07-DomainCutover/) use
- An accepted domain in the destination for new group addresses, e.g. a staging domain. Source domains get added at cutover

## Usage

```powershell
# 1. Source tenant: export. Changes nothing.
.\Copy-CollaborationMap.ps1 -Export -TenantId <source-tenant-guid> -SharePointAdminUrl https://fabrikam-admin.sharepoint.com

# 2. Destination tenant: preview, then import
Disconnect-ExchangeOnline -Confirm:$false
.\Copy-CollaborationMap.ps1 -Import -TenantId <destination-tenant-guid> -ExportFolder .\Collaboration-fabrikam.com-<date> `
    -MappingFile .\UserMapping.csv -AddressDomain contoso.com -EnableTeams -WhatIf
.\Copy-CollaborationMap.ps1 -Import -TenantId <destination-tenant-guid> -ExportFolder .\Collaboration-fabrikam.com-<date> `
    -MappingFile .\UserMapping.csv -AddressDomain contoso.com -EnableTeams

# Destination already has a group called 'sales'?
.\Copy-CollaborationMap.ps1 -Import ... -AliasPrefix fab-

# One team at a time
.\Copy-CollaborationMap.ps1 -Import ... -Group sales@fabrikam.com

# Name the account (Exchange Online, and Graph with -EnableTeams)
.\Copy-CollaborationMap.ps1 -Import -TenantId <destination-tenant-guid> -AdminUpn admin@contoso.onmicrosoft.com -ExportFolder ... -MappingFile ... -AddressDomain contoso.com
```

**Tip: name the account.** With both tenants open on one PC, `Connect-ExchangeOnline` and `Connect-MgGraph` happily reuse whichever session is already there. Add `-AdminUpn admin@contoso.onmicrosoft.com` and the script disconnects any session for another account or tenant, signs in to Exchange Online (and Graph, with `-EnableTeams`) as the one you named, and prints `Signed in as X to tenant Y` before it changes anything. On the live cutover this would have saved a confusing error at the worst moment. The signed-in account becomes the first owner of every new group, which is one more reason to make sure it's the right one.

Then, in your copy tool, copy each Team's channels and each site's libraries into the groups that now exist, rather than letting it create new ones.

## Output

- **Export:** a `Collaboration-<domain>-<date>` folder with `Groups.csv`, `GroupLinks.csv` and, with `-SharePointAdminUrl`, `Sites.csv`.
- **Import:** `Collaboration-Import-<domain>-<date>.csv`, with every action marked Created, Added, Exists, Skipped, Problem or WhatIf, and why.

The CSVs contain names and addresses. Keep them out of source control and shared drives. The repo's `.gitignore` already excludes `*.csv`.

## Lessons learnt

- **Create the groups yourself before you copy.** That's the whole point of this script. The address lands on the domain you chose, and the site URL is what you want rather than what the tool picked. In my experience the site URL follows the alias, so choose the alias (and any `-AliasPrefix`) with the URL in mind.
- **Private and shared channels aren't created by this.** They have their own sites and their own membership. `Sites.csv` lists them so you know they're there. Check what your copy tool does with them, and plan to create any it doesn't handle by hand.
- **Import the user mapping into the copy tool before you start.** I forgot once, and the content came across without the members. The import here adds owners and members up front, but the copy tool still needs the mapping for permissions and authorship inside the content.
- **Shared mailboxes can be group members.** In the source I found shared mailboxes sitting in group membership. If they're in your mapping file they come across like anyone else. If they're not, the report lists them as skipped, so check whether they matter.
- **You end up owning everything.** The signed-in admin is the first owner of every group the script creates, so there's always an owner while the real ones are added. Remove yourself with `Remove-UnifiedGroupLinks` once the copy is done, if you don't want to stay.
- **Teams can take a moment.** Graph sometimes can't see a brand-new group straight away. The script retries a few times. If it still fails, the group is fine; re-run later and it'll only do the Teams part.
- **Welcome messages are off for a reason.** Otherwise every member gets a 'you've been added to' email from a tenant they haven't started using yet, weeks before cutover.

## Three ways to do it

| Route | How |
| --- | --- |
| Paid tool | ShareGate can create the Team for you while copying, but the address and site URL are its choice, and a name clash gives you a "2" on the end. Create the groups with this script, then use "copy channels into an existing team" |
| Microsoft's own | Teams admin centre or `New-UnifiedGroup` and `New-Team` by hand, then add owners and members. Native cross-tenant migration doesn't move shared Teams content at all |
| By hand | Create each Team in the Teams client with the right name, then add members. Fine for four Teams. Check the email address it was given before you copy anything into it |

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 1.1.0 | 2026-11-10 | `-AdminUpn`: disconnects Exchange Online and Graph sessions for any other account or tenant, signs in as the one named, prints who and where |
| 1.0.0 | 2026-11-09 | First release |
