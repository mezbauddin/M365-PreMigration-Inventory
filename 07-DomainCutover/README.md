# Domain cutover

**TL;DR:** Four scripts for the night a domain moves between tenants. Record the DNS, release the domain from the source, add and verify it in the destination, then give everyone their old addresses back, with X500 addresses so replies to old mail don't bounce. Every step supports `-WhatIf`, and every step checks it's in the right tenant before changing anything.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.0.0 |
| Released | 9 November 2026 |
| Licence | MIT |

Part 7 of my [M365 tenant-to-tenant migration toolkit](../README.md).

## Why you need this

A domain can only be verified in one Microsoft 365 tenant at a time. So cutover is a relay. The source has to let go of every address, sign-in name and group on the domain before the destination can claim it. Then DNS has to point at the destination. Then every migrated mailbox has to answer to its old address again.

## What happens without it

- **The source won't release the domain.** One leftover alias on a group, a guest, or a deleted user still in the recycle bin is enough to block removal. You find out at 9pm on cutover night.
- **Mail bounces in the gap.** Between the domain leaving the source and MX pointing at the destination, inbound mail can bounce. High DNS TTLs stretch that gap from minutes to a day.
- **Replies to old mail bounce.** Migrated messages and Outlook's autocomplete address people by their old internal Exchange address (LegacyExchangeDN), not their email address. Without a matching X500 address, those replies fail with IMCEAEX non-delivery reports, sometimes for months.
- **People sign in with the wrong name.** Staged accounts sit on a temporary domain until someone switches them.

## The scripts, in order

| # | Script | Tenant | What it does | Changes anything? |
| --- | --- | --- | --- | --- |
| 1 | `Export-DomainDns.ps1` | Neither (DNS) | Records MX, SPF, autodiscover, DKIM, DMARC and the rest, with TTLs. Flags TTLs to lower | No |
| 2 | `Remove-SourceDomain.ps1 -SnapshotOnly` | Source | Snapshots every recipient's addresses and LegacyExchangeDN | No, writes a CSV |
| 3 | `Remove-SourceDomain.ps1 -RemoveDomain` | Source | Moves the default domain, sign-in names, primary addresses and aliases off the domain, lists blockers, removes it | Yes, source |
| 4 | `Add-TargetDomain.ps1`, then `-Verify`, then `-EnableDkim` | Destination | Adds the domain, gives you the TXT to verify, then prints the MX, SPF, autodiscover and DKIM records read from the tenant | Yes, destination |
| 5 | `Switch-TargetAddresses.ps1` | Destination | Adds old addresses and X500s, makes them primary and the sign-in name for migrated recipients, optional temporary passwords | Yes, destination |

## A cutover night, end to end

**Days before**

```powershell
.\Export-DomainDns.ps1 -Domain fabrikam.com -Server 1.1.1.1      # lower anything flagged
.\Remove-SourceDomain.ps1 -TenantId <source-guid> -Domain fabrikam.com -SnapshotOnly
.\Remove-SourceDomain.ps1 -TenantId <source-guid> -Domain fabrikam.com -WhatIf   # read the blockers
```

**On the night**

```powershell
# 0. Final incremental copy. Save today's DNS as the rollback copy
.\Export-DomainDns.ps1 -Domain fabrikam.com -Server 1.1.1.1

# 1. Source: snapshot again, release, remove
.\Remove-SourceDomain.ps1 -TenantId <source-guid> -Domain fabrikam.com -RemoveDomain

# 2. Destination: add, then put the TXT record in DNS
Disconnect-ExchangeOnline -Confirm:$false; Disconnect-MgGraph
.\Add-TargetDomain.ps1 -TenantId <dest-guid> -Domain fabrikam.com

# 3. Once the source has finished removing it and the TXT resolves
.\Add-TargetDomain.ps1 -TenantId <dest-guid> -Domain fabrikam.com -Verify
#    Change MX, SPF, autodiscover and the DKIM CNAMEs to what it prints

# 4. Addresses, X500s, sign-in names, passwords
.\Switch-TargetAddresses.ps1 -TenantId <dest-guid> -MappingFile .\UserMapping.csv `
    -SourceSnapshot .\SourceRecipients-<date>.csv -Domain fabrikam.com -WhatIf
.\Switch-TargetAddresses.ps1 -TenantId <dest-guid> -MappingFile .\UserMapping.csv `
    -SourceSnapshot .\SourceRecipients-<date>.csv -Domain fabrikam.com -ResetPasswords

# 5. When the DKIM CNAMEs resolve
.\Add-TargetDomain.ps1 -TenantId <dest-guid> -Domain fabrikam.com -EnableDkim
```

`Remove-SourceDomain` asks you to confirm each change, because it's the one that can't easily be undone. Add `-Confirm:$false` once you've read the `-WhatIf` output and you're happy.

## How Switch-TargetAddresses decides

| Recipient | Old addresses | X500 | Primary address | Sign-in name |
| --- | --- | --- | --- | --- |
| In `TargetBuildPlan.csv` (built by part 2) | Added | Added | Old address | Old address |
| List created by [Copy-MailConfig](../05-MailConfig/) (tagged) | Added | Added | Old address | n/a |
| Anyone else in the mapping file, e.g. people already in the destination, Microsoft 365 groups | Added as aliases | Added | Unchanged | Unchanged |
| Mail contacts | No | Added | Unchanged | n/a |

## Requirements

- `ExchangeOnlineManagement` 3.x, `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users`, `Microsoft.Graph.Identity.DirectoryManagement`
- Global Administrator in each tenant
- Directory sync off in the source. Synced users' sign-in names and addresses can't be changed in the cloud
- Windows for `Export-DomainDns.ps1` (`Resolve-DnsName`)

## Lessons learnt

- **Snapshot first, today.** Once the source has released the domain, the old addresses are gone from it. The snapshot is the only record of who had which address and LegacyExchangeDN. The X500 step depends on it.
- **Deleted users block removal.** A user in the recycle bin keeps their sign-in name on the domain. Purge them, or restore and rename them, before cutover night.
- **The default domain can't be removed.** If the moving domain is the source's default, switch the default to onmicrosoft.com first. The script does.
- **Removal isn't instant.** The source can say "removed" while the destination still says "already in use" for a while. Plan for a gap, and set expectations that mail may queue at senders.
- **Read the DNS from the tenant, don't copy it from a blog.** Newer domains can get a different MX format. `Add-TargetDomain` prints what the destination actually wants.
- **Keep third-party senders in SPF.** If the old SPF record included a CRM or a marketing platform, carry those includes over.
- **Sign-in changes have side effects.** OneDrive URLs rename themselves, Office and Teams ask people to sign in again, and phones need their accounts re-adding. Warn people the day before.
- **Two tenants, one PC.** Disconnect between the source and destination steps. Every script here checks the tenant ID, but you'll be running plenty of commands by hand too.

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 1.0.0 | 2026-11-09 | First release: Export-DomainDns, Remove-SourceDomain, Add-TargetDomain, Switch-TargetAddresses |

## Deleted users blocking the domain: `Clear-DeletedUserDomain.ps1`

If `Remove-SourceDomain -WhatIf` lists **Blocker: Deleted user**, those are leavers in the recycle bin still holding a sign-in on the domain. Microsoft won't release the domain, and won't rename a deleted user. This restores each one as `<name>-old@<tenant>.onmicrosoft.com`, blocks it and deletes it again, so nothing is purged and everything stays recoverable for 30 more days.

```powershell
.\Clear-DeletedUserDomain.ps1 -TenantId <source-guid> -Domain fabrikam.com -WhatIf
.\Clear-DeletedUserDomain.ps1 -TenantId <source-guid> -Domain fabrikam.com
```

Run it the day before cutover, not on the night: it's one more thing that can surprise you.
