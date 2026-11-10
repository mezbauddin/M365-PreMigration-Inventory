# Domain cutover

**TL;DR:** Eight scripts for the night a domain moves between Microsoft 365 tenants. Check you're ready, record the DNS, get the destination's verification record published first, release the domain from the source, verify it in the destination within minutes, then give everyone their old addresses back with X500s, and block the old accounts. Every script that changes something supports `-WhatIf`, signs in as the account you name, and refuses to run against the wrong tenant. These are field notes from a live cutover on 10 November 2026, written so you don't repeat my evening.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.1.0 |
| Released | 10 November 2026 |
| Licence | MIT |

Part 7 of my [M365 tenant-to-tenant migration toolkit](../README.md). Examples use fabrikam.com (the domain moving, source tenant fabrikam.onmicrosoft.com) and the contoso destination tenant.

## Why this part matters

A domain can only be verified in one Microsoft 365 tenant at a time. So cutover is a relay: the source lets go of every address, sign-in name, app and leftover reference on the domain, the destination claims it, and every migrated mailbox answers to its old address again. Mail to the domain bounces from the moment the source strips the addresses until the destination verifies it. The whole craft of cutover night is making that gap as short as possible and having nothing surprise you inside it.

## Order of play

This is the add-first order. It turned a potential multi-hour bounce window into a few minutes, because verification is instant once the source releases the domain and the TXT record is already live.

| # | When | Script | Tenant | What it does | Changes anything? |
| --- | --- | --- | --- | --- | --- |
| 1 | T-3 days and T-1 hour | `Test-CutoverReadiness.ps1` | Both, read-only | Go/no-go: blockers, snapshot, mapping, conflicts, domain state, DNS, Conditional Access | No |
| 2 | T-3 days | `Export-DomainDns.ps1 -FullScan` | DNS | Finds where DNS really lives, NS TTL, and every record it can see | No |
| 3 | At least 48 h before, only if you can't log in to the DNS host | `Set-DomainDnsRecords.ps1 -FromScan`, then move the name servers | DNS | Replicates the zone somewhere you control | DNS |
| 4 | T-1 day | `Remove-SourceDomain.ps1 -SnapshotOnly` and `-BlockersOnly` | Source | The `-pre` snapshot, and the list of blockers to clear | Writes CSVs |
| 5 | T-1 day | `Clear-DeletedUserDomain.ps1` | Source | Frees the domain from deleted users | Source |
| 6 | Night, step 1 | `Add-TargetDomain.ps1` | Destination | Adds the domain unverified, gives you the MS=ms TXT | Destination |
| 7 | Night, step 2 | Publish the TXT, then `Add-TargetDomain.ps1 -WaitForDns` | DNS | Confirms every authoritative name server serves the TXT | DNS |
| 8 | Night, step 3 | `Remove-SourceDomain.ps1 -RemoveDomain` | Source | Moves everything off and removes the domain. Mail starts bouncing here | Source |
| 9 | Night, step 4, straight after | `Add-TargetDomain.ps1 -Verify` | Destination | Verifies. Mail stops bouncing once MX and the tenant agree | Destination |
| 10 | Night, step 5 | `Switch-TargetAddresses.ps1` with the `-pre` snapshot | Destination | Old addresses, X500s, primary addresses, sign-in names, passwords | Destination |
| 11 | Night, step 6 | DKIM CNAMEs, then `Add-TargetDomain.ps1 -EnableDkim` | DNS, destination | Signs outbound mail as the domain again | Destination |
| 12 | Next morning | `Block-SourceSignIn.ps1` | Source | Stops people working in the old tenant | Source |

```powershell
# T-3 days
.\Test-CutoverReadiness.ps1 -SourceTenantId <source-guid> -DestinationTenantId <destination-guid> `
    -SourceAdmin admin@fabrikam.onmicrosoft.com -DestinationAdmin admin@contoso.onmicrosoft.com `
    -Domain fabrikam.com -MappingFile .\UserMapping.csv -SourceSnapshot .\SourceRecipients-pre-<stamp>.csv
.\Export-DomainDns.ps1 -Domain fabrikam.com -FullScan

# T-1 day
.\Remove-SourceDomain.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -Domain fabrikam.com -SnapshotOnly
.\Remove-SourceDomain.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -Domain fabrikam.com -BlockersOnly
.\Clear-DeletedUserDomain.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -Domain fabrikam.com

# The night
.\Add-TargetDomain.ps1 -TenantId <destination-guid> -AdminUpn admin@contoso.onmicrosoft.com -Domain fabrikam.com
#   publish the MS=ms TXT it prints, then:
.\Add-TargetDomain.ps1 -TenantId <destination-guid> -AdminUpn admin@contoso.onmicrosoft.com -Domain fabrikam.com -WaitForDns
.\Remove-SourceDomain.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -Domain fabrikam.com -FixApplications -RemoveDomain
.\Add-TargetDomain.ps1 -TenantId <destination-guid> -AdminUpn admin@contoso.onmicrosoft.com -Domain fabrikam.com -Verify
.\Switch-TargetAddresses.ps1 -TenantId <destination-guid> -AdminUpn admin@contoso.onmicrosoft.com -MappingFile .\UserMapping.csv `
    -SourceSnapshot .\SourceRecipients-pre-<stamp>.csv -Domain fabrikam.com -WhatIf
.\Switch-TargetAddresses.ps1 -TenantId <destination-guid> -AdminUpn admin@contoso.onmicrosoft.com -MappingFile .\UserMapping.csv `
    -SourceSnapshot .\SourceRecipients-pre-<stamp>.csv -Domain fabrikam.com -RemoveConflictingGuests -ResetPasswords
#   change the two DKIM CNAMEs, then:
.\Add-TargetDomain.ps1 -TenantId <destination-guid> -AdminUpn admin@contoso.onmicrosoft.com -Domain fabrikam.com -EnableDkim

# Next morning
.\Block-SourceSignIn.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -MappingFile .\UserMapping.csv
```

Tip: keep two PowerShell windows, one per tenant, and say the tenant out loud before each command. The scripts check, but your hand-typed commands don't.

## What changes in DNS

Less than most people expect. Microsoft 365 builds the MX host from the domain name, not the tenant, so it's the same in both tenants.

| Record | Example | Action at cutover |
| --- | --- | --- |
| MX | `@ MX 0 fabrikam-com.mail.protection.outlook.com` | **Unchanged** |
| SPF | `@ TXT v=spf1 include:spf.protection.outlook.com -all` | **Unchanged** (keep any third-party includes) |
| Autodiscover | `autodiscover CNAME autodiscover.outlook.com` | **Unchanged** |
| Device registration and enrolment | `enterpriseregistration`, `enterpriseenrollment` CNAMEs | **Unchanged** |
| Verification | `@ TXT MS=ms<number>` | **Change**: add the destination's value. It sits alongside SPF and the old one |
| DKIM | `selector1._domainkey`, `selector2._domainkey` CNAMEs | **Change**: they point at the tenant. Swap after verification, then `-EnableDkim` |
| Skype for Business legacy | `sip`, `lyncdiscover`, `_sip._tls`, `_sipfederationtls._tcp`, `msoid` | **Skip**. Teams doesn't need them |
| Apex CNAME | `@ CNAME contoso.sharepoint.com` | **Never apply**. An apex CNAME breaks MX and TXT for the whole domain. The admin centre wizard sometimes suggests it |

`Add-TargetDomain` puts each record in one of these buckets for you (Unchanged, Change, Add, Skip, DoNotApply, Check) by comparing against the authoritative name server.

## The scripts

Each section: what it does, why, what happens without it, and three ways to do the job: a paid tool, Microsoft's own portals, and by hand or with this script.

### 1. `Test-CutoverReadiness.ps1` (1.0.0): the go/no-go

**What:** read-only checks across both tenants and DNS, printed as PASS, WARN or FAIL, saved to `CutoverReadiness-<stamp>.csv`. GO means no FAIL. It signs in to the source then the destination with the accounts you name, and runs `..\tools\Test-ModuleHealth.ps1` first if it's there.

**Why:** every check in it is something that stopped or slowed a real cutover: a leaver in the recycle bin, an app registration nobody knew about, a guest holding someone's old address, a TXT record on one name server but not the other, a Conditional Access policy that blocks the admin's home sign-in.

**Without it:** you find each of those at 9pm, one at a time.

- **Paid tool:** migration tools have pre-flight reports for the data copy. None I've used checks domain blockers or destination address conflicts.
- **Microsoft native:** Entra admin centre (Deleted users, App registrations), Exchange admin centre (Recipients, search each old address), the Microsoft 365 admin centre domain page. Allow an hour.
- **By hand:** `nslookup -type=txt fabrikam.com <each name server>` and the Graph queries in the script.

Tip: run it at T-3 days, fix everything, then again an hour before. Read every WARN.

### 2. `Export-DomainDns.ps1` (1.1.0): know your DNS

**What:** the default run records the Microsoft 365 records with TTLs and flags any above `-MaxTtl`. `-FullScan` finds the authoritative name servers, asks one of them directly for the apex records and about 100 common names (plus your `-Wordlist`), and writes `Domain, Name, Type, TTL, Value`. It prints where DNS is hosted, warns when the NS TTL is a day or more, and warns on apex A records with a 60-second TTL.

**Why:** DNS is often not where you think. On the live cutover the registrar was obvious, but the name servers pointed at a cloud DNS service that nobody could log in to. Changing records at the registrar did nothing.

**Without it:** you change the wrong place, or you lose a record nobody remembered when you move DNS.

- **Paid tool:** your DNS host's export, or a commercial DNS scanner.
- **Microsoft native:** none. The Microsoft 365 admin centre shows only the records it wants.
- **By hand:** `Resolve-DnsName fabrikam.com -Type NS`, then query that server for each name.

Trap: a wordlist scan only finds names it asks for. Zone transfers are nearly always refused. Ask the website owner and every SaaS that sends mail as the domain for their records.
Trap: apex A records with a 60-second TTL in CDN ranges are probably an alias record. Copying them as static A records works today and breaks when the CDN moves.
Trap: NS records often have a 172800-second (48 hour) TTL. Move DNS at least that long before cutover, never on the night.
Tip: `Resolve-DnsName` doesn't do CAA. Check CAA in a web DNS lookup if you issue certificates.

```powershell
.\Export-DomainDns.ps1 -Domain fabrikam.com -Server 1.1.1.1
.\Export-DomainDns.ps1 -Domain fabrikam.com -FullScan -Wordlist .\extra-names.txt
```

### 3. `Set-DomainDnsRecords.ps1` (1.0.0): write DNS through the API

**What:** takes a records CSV (`Type, Name, Value, TTL, Priority`) or the `-FullScan` CSV (`-FromScan`) and makes each Type+Name group in a GoDaddy or Cloudflare zone match it. Backs up the zone to JSON first. Never touches NS or SOA. Refuses any apex CNAME. Lists, but leaves alone, records in the zone that aren't in your file. `-WhatIf` shows the plan.

**Why:** two uses. Replicating a zone you can't log in to into one you can, days ahead. And applying the Change and Add rows on the night without clicking through a portal at speed.

**Without it:** fifty records typed by hand, at night, with mail down. One typo in an SPF record and mail to big providers starts going to junk.

- **Paid tool:** DNS management platforms with import from file.
- **Microsoft native:** if you host DNS in Microsoft 365 (rare for business domains), the admin centre can manage the records for you.
- **By hand:** the provider's DNS page. Fine for five records.

How it writes: GoDaddy uses `GET /v1/domains/{domain}/records` for the backup and `PUT /v1/domains/{domain}/records/{type}/{name}` per group, which also overwrites GoDaddy's default Parked A record. Cloudflare looks up the zone ID, then deletes and creates per group. Credentials come from `GODADDY_KEY`/`GODADDY_SECRET` or `CLOUDFLARE_API_TOKEN`, or a prompt. Nothing is stored.

Trap: **a GoDaddy API key works on every domain in the account. Delete it in the GoDaddy developer portal as soon as you're done.**
Trap: GoDaddy adds its own `_domainconnect` CNAME. Leave it.
Tip: for Cloudflare, make a token with DNS edit on that one zone only.

```powershell
.\Set-DomainDnsRecords.ps1 -Domain fabrikam.com -Provider GoDaddy -FromScan .\DomainDns-FullScan-<stamp>.csv -WhatIf
.\Set-DomainDnsRecords.ps1 -Domain fabrikam.com -Provider Cloudflare -RecordsFile .\cutover-records.csv
```

### 4. `Remove-SourceDomain.ps1` (1.1.0): release the domain

**What:** snapshots every recipient (addresses and LegacyExchangeDN) to `SourceRecipients-pre-<stamp>.csv` with `SnapshotStage=BeforeRelease`, lists blockers, moves sign-in names, primary addresses and aliases to onmicrosoft.com, and with `-RemoveDomain` removes the domain. Blocker classes:

| Blocker | What it is | Fix |
| --- | --- | --- |
| Deleted user | A leaver in the recycle bin still holds a sign-in on the domain | `Clear-DeletedUserDomain.ps1` |
| Synced user | Directory sync still owns the account | Turn sync off first |
| Application | An app registration has an identifier URI on the domain | `-FixApplications` moves it to `https://<tenant>.onmicrosoft.com/<path>` |
| ProxyOnly | A user with no mailbox whose Entra proxyAddresses still hold the domain. Graph can't edit them | Admin centre: Settings > Domains > the domain > Remove domain wizard |
| SyncLag | Exchange won't drop an address because it is still the WindowsLiveId after the sign-in change | Waits and retries every minute for `-RetryMinutes` (default 10) |
| Reference | Anything else the directory links to the domain | By hand, or the admin centre wizard |

Blockers are listed in `-WhatIf` and `-BlockersOnly` runs too. On a real run, deleted users, synced users and apps stop the script before it changes anything. The last line tells you what to do next: Add-TargetDomain -Verify in the destination.

**Why:** Microsoft won't remove a domain while anything references it, and the error tells you very little.

**Without it:** you strip everyone's addresses (mail starts bouncing), then find a blocker you can't clear quickly.

- **Paid tool:** some migration suites (Quest On Demand Migration, for example) have a domain move feature that does this and the destination side together. ShareGate doesn't touch domains.
- **Microsoft native:** Microsoft 365 admin centre > Settings > Domains > Remove domain. The wizard moves users, groups and proxy addresses. It doesn't snapshot anything and doesn't fix app registrations, but it is the only way to fix ProxyOnly users.
- **By hand:** `Set-Mailbox -EmailAddresses @{Remove=...}` and `Update-MgUser -UserPrincipalName` per object. Fine for ten.

Trap: a snapshot taken after the addresses are stripped is useless. If no recipient holds the domain when it snapshots, the file is named `-post` and the script warns in red. Use the `-pre` file with Switch-TargetAddresses.
Trap: `SyncLag` is normal for a few minutes after the sign-in change. Don't panic; let the retry run.

```powershell
.\Remove-SourceDomain.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -Domain fabrikam.com -BlockersOnly
.\Remove-SourceDomain.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -Domain fabrikam.com -FixApplications -RemoveDomain
```

### 5. `Clear-DeletedUserDomain.ps1` (1.1.0): deleted users holding the domain

**What:** restores each deleted user that holds the domain as `<name>-old@<tenant>.onmicrosoft.com`, blocks it, deletes it again. Nothing is purged; every account stays recoverable for 30 more days.

**Why:** Microsoft won't remove a domain while a deleted user holds it, and won't rename a deleted user.

**Without it:** you purge them (permanent) or restore and rename them one at a time in the portal.

- **Paid tool:** none needed.
- **Microsoft native:** Entra admin centre > Users > Deleted users > restore, rename in Properties, block, delete.
- **By hand:** `Remove-MgDirectoryDeletedItem` to purge, if nobody will ever need them.

Trap: leavers who held an admin role can only be restored by a Global Administrator or Privileged Role Administrator.

```powershell
.\Clear-DeletedUserDomain.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -Domain fabrikam.com -WhatIf
```

### 6. `Add-TargetDomain.ps1` (1.1.0): add, wait, verify, DKIM

**What:** adds the domain unverified and prints the MS=ms TXT. `-WaitForDns` polls every authoritative name server every 30 seconds until all serve it (up to `-WaitMinutes`, default 30), then carries on to `-Verify` if you gave it. After verification it prints every record with an Action (see the DNS table) and creates DKIM keys. `-EnableDkim` once the CNAMEs resolve.

**Why:** add first, and the destination's TXT is already live everywhere when the source lets go. Verification was instant on the live cutover.

**Without it:** you add the domain after removing it from the source, wait for DNS, and mail bounces the whole time.

- **Paid tool:** domain move features in migration suites do this step too.
- **Microsoft native:** Microsoft 365 admin centre > Settings > Domains > Add domain. Ignore its Skype records and any apex SharePoint CNAME.
- **By hand:** `New-MgDomain`, `Get-MgDomainVerificationDnsRecord`, `Confirm-MgDomain`, `New-DkimSigningConfig`.

Trap: a TXT on one name server and not the other means verification fails half the time. That's why `-WaitForDns` checks all of them, not your local resolver.
Trap: verified in Entra doesn't mean Exchange knows yet. Wait a few minutes before DKIM.

### 7. `Switch-TargetAddresses.ps1` (1.1.0): old addresses and X500s

**What:** for each mapping row (and each list Copy-MailConfig created), adds the old addresses and an X500 from the old LegacyExchangeDN, and for migrated recipients makes the old address primary and the sign-in name. Before anything, it checks the snapshot is the `-pre` one (addresses on the domain and LegacyExchangeDN values present, at least 20% of mapped sources found) and stops otherwise; `-Force` overrides. It checks each old address against the destination: if a guest, mail user, contact or another mailbox holds it, that address is logged as `Conflict` and skipped. `-RemoveConflictingGuests` deletes conflicting guests (recycle bin) and retries.

**Why:** migrated mail and Outlook autocomplete address people by their old internal Exchange address. Without the X500, replies bounce with IMCEAEX reports for months.

**Without it:** the night "works", and the helpdesk gets bounce reports for weeks.

- **Paid tool:** most migration tools can add X500s in the final pass if you give them the LegacyExchangeDN. Check they do groups and contacts too.
- **Microsoft native:** Exchange admin centre > the mailbox > Email addresses: add the SMTP and an X500 address by hand.
- **By hand:** `Set-Mailbox -EmailAddresses @{Add='smtp:...','X500:...'}`.

Trap: destination guest accounts often hold people's old addresses (they were invited before the move). The add fails with "The proxy address ... is already being used by ...".
Trap: if the destination hosts other companies, there may be same-name mailboxes on other domains. Matching is strictly by the mapping file's Destination. Never by name.

| Recipient | Old addresses | X500 | Primary address | Sign-in name |
| --- | --- | --- | --- | --- |
| In `TargetBuildPlan.csv` (part 2) | Added | Added | Old address | Old address |
| List created by [Copy-MailConfig](../05-MailConfig/) (tagged) | Added | Added | Old address | n/a |
| Anyone else in the mapping file | Added as aliases | Added | Unchanged | Unchanged |
| Mail contacts | No | Added | Unchanged | n/a |

### 8. `Block-SourceSignIn.ps1` (1.0.0): close the old door

**What:** for each source account (by UPN, mail, proxy address, or `<name>@<tenant>.onmicrosoft.com` after the release), sets accountEnabled to false and reads it back. Only says Blocked when the read-back agrees. On 403 Authorization_RequestDenied it checks directory roles and tells you to block admins in the admin centre (https://admin.microsoft.com/#/users). Skips your own admin account and anything in `-Exclude`.

**Why:** phones and saved sessions keep people working in the old tenant, and that mail never reaches the new one.

**Without it:** a week later someone asks why nobody replied to them, and the answer is sitting in a source mailbox.

- **Paid tool:** not needed.
- **Microsoft native:** Microsoft 365 admin centre > Active users > select > Block sign-in.
- **By hand:** `Update-MgUser -AccountEnabled:$false`, then check it.

Trap: blocking doesn't end live sessions. Revoke them too (`Revoke-MgUserSignInSession`) if it must be immediate.
Tip: keep one break-glass Global Administrator in the source unblocked until it is decommissioned.

```powershell
.\Block-SourceSignIn.ps1 -TenantId <source-guid> -AdminUpn admin@fabrikam.onmicrosoft.com -MappingFile .\UserMapping.csv -WhatIf
```

## Lessons learned on a live cutover

- **Sessions lie.** The scripts reused whatever Exchange or Graph session was open and twice landed in the wrong tenant. The tenant ID check stopped them both times. Now every script takes `-AdminUpn`, disconnects anything that doesn't match, and prints who it signed in as before changing anything.
- **Add first.** Add the domain to the destination before releasing it, publish the TXT, confirm both name servers serve it, then release, then verify. Verification was instant.
- **Blockers hide.** Beyond deleted users: app registrations with identifier URIs on the domain, users with no mailbox whose Entra proxy addresses still held it (only the admin centre wizard fixed those), and addresses Exchange wouldn't drop for a few minutes because they were still the WindowsLiveId.
- **Mail bounces in the gap.** From the moment addresses are stripped until the destination verifies. Keep the gap short and tell senders in advance.
- **DNS lives where the NS records say.** It was a cloud DNS service nobody could log in to. Scan it, replicate it somewhere you control, and move the name servers at least 48 hours ahead (the NS TTL).
- **Fast apex TTLs mean alias records.** 60-second A records in CDN ranges were an alias. Copying them as static A records is a risk.
- **Snapshot before, not after.** One run used a snapshot taken after the addresses were stripped. Nearly everyone logged "Not in the source snapshot" and got no aliases or X500s. The scripts now name the file `-pre` or `-post` and refuse a `-post` snapshot.
- **Guests hold old addresses.** Destination guest accounts had users' old addresses as proxies, so alias adds failed. The destination also hosted other companies with same-name mailboxes. Match by the mapping file only.
- **Most DNS doesn't change.** MX, SPF and autodiscover are identical between tenants. Only the verification TXT and DKIM CNAMEs change. Skip the Skype records and never apply an apex CNAME to sharepoint.com.
- **Block the old accounts, and check it worked.** Accounts with admin roles return 403 through Graph and need the admin centre. A loop printed "Blocked" even when the update failed: always try/catch and read back.
- **APIs beat portals at night.** The GoDaddy API (back up with GET, then PUT per type and name) kept NS and SOA untouched and replaced the Parked record cleanly. Then delete the key.

## Requirements

- `ExchangeOnlineManagement` 3.x, `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users`, `Microsoft.Graph.Applications`, `Microsoft.Graph.Identity.DirectoryManagement`
- Global Administrator in each tenant for the changing steps; Global Reader is enough for Test-CutoverReadiness
- Directory sync off in the source
- Windows for `Resolve-DnsName` (Export-DomainDns, `-WaitForDns`, the Action column, DNS checks)
- A GoDaddy API key and secret or a Cloudflare API token for Set-DomainDnsRecords

## Changelog

| Script | Version | Date | Change |
| --- | --- | --- | --- |
| Test-CutoverReadiness.ps1 | 1.0.0 | 2026-11-10 | New: go/no-go across both tenants and DNS |
| Set-DomainDnsRecords.ps1 | 1.0.0 | 2026-11-10 | New: GoDaddy and Cloudflare zone writes with backup, from a records file or a full scan |
| Block-SourceSignIn.ps1 | 1.0.0 | 2026-11-10 | New: block source accounts with read-back and admin role detection |
| Export-DomainDns.ps1 | 1.1.0 | 2026-11-10 | `-FullScan`, `-Wordlist`, DNS host, NS TTL and alias warnings |
| Remove-SourceDomain.ps1 | 1.1.0 | 2026-11-10 | `-AdminUpn`, `-BlockersOnly`, `-FixApplications`, `-RetryMinutes`; Application, ProxyOnly and SyncLag blockers; blockers shown under `-WhatIf`; stops before changes on hard blockers; `-pre`/`-post` snapshot names with SnapshotStage |
| Add-TargetDomain.ps1 | 1.1.0 | 2026-11-10 | `-AdminUpn`, `-WaitForDns`, `-WaitMinutes`; Action column against the authoritative server; add-first order |
| Switch-TargetAddresses.ps1 | 1.1.0 | 2026-11-10 | `-AdminUpn`, `-Force`, `-RemoveConflictingGuests`; snapshot sanity check; conflict pre-scan |
| Clear-DeletedUserDomain.ps1 | 1.1.0 | 2026-11-10 | `-AdminUpn` and session check |
| All four original scripts | 1.0.x | 2026-11-09 | First release |
