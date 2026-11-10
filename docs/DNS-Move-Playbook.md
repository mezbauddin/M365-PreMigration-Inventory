# DNS move playbook

**TL;DR:** If you can't log in to the place the domain's DNS is actually hosted, you can't cut over. Move DNS hosting first, at least 48 hours before the cutover: scan every record from public DNS, recreate them at the new provider (usually the registrar), check them against the new name servers and a public resolver, then switch the name servers. Keep the old name server list as your rollback, watch for alias records that keep the website up, and delete the API keys when you're done.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.0.0 |
| Released | 10 November 2026 |
| Licence | MIT |

Part of the [docs](README.md) for my [M365 tenant-to-tenant migration toolkit](../README.md).

## Why you need it

Cutover night needs four DNS changes in quick succession: a verification TXT, then MX, SPF, autodiscover and DKIM. Whoever does them must be able to log in to the **DNS host**. That's often not the registrar. On the live migration the domain was registered at one company but its name servers pointed at a cloud DNS service in an account nobody could reach: an old IT provider had set it up years before. The registrar login worked, and was useless.

## What happens without it

- **You find out on the night.** The source releases the domain, you go to add the TXT, and you can't. Mail bounces until someone finds the old provider's password.
- **You rebuild DNS in a hurry** from memory and a screenshot, miss a record, and the website, a VPN or a supplier's integration stops working.
- **You switch name servers too late.** Name server changes are cached for up to 48 hours. Switch on the day and half the internet still asks the old host.

## When to do it

At **T-14**, confirm who can log in to the DNS host (`Resolve-DnsName fabrikam.com -Type NS` shows where it is). If it isn't someone on the project, move DNS hosting **at least 48 hours before cutover**, ideally a week. Never on the same day.

## The steps

| # | Step | How | Toolkit |
| --- | --- | --- | --- |
| 1 | Find where DNS lives | `Resolve-DnsName fabrikam.com -Type NS`. Compare with the registrar's records. If they differ, the NS answer wins | |
| 2 | Full scan | Read every record public DNS will give you: apex, www, MX, SPF and other TXT, autodiscover, DKIM selectors, DMARC, SRV, common and discovered subdomains, with TTLs | [07 Export-DomainDns.ps1 -FullScan](../07-DomainCutover/) |
| 3 | Fill the gaps | A scan can't list records nobody asks for. Ask the business and old providers about VPN, remote desktop, CRM, marketing and verification records. Add names to a wordlist and scan again | `-Wordlist .\extra-names.txt` |
| 4 | Replicate | Create every record at the new provider, with the same values. Do it through the provider's API so nothing is typed by hand | [07 Set-DomainDnsRecords.ps1 -FromScan](../07-DomainCutover/) `-WhatIf` first |
| 5 | Verify against the new name servers | Query the new provider's name servers directly, before switching. Every record should answer the same as the old host | `Resolve-DnsName <name> -Type <type> -Server <new NS>` |
| 6 | Lower TTLs | At the old host if you can, and at the new one: 300 to 3600 seconds on MX, autodiscover, SPF and anything you'll change at cutover | |
| 7 | Save the rollback | Write down the old name server list exactly, and keep the full scan CSV | The scan CSV |
| 8 | Switch NS | At the registrar, replace the name servers with the new provider's | Registrar portal |
| 9 | Verify against a public resolver | Over the next hours, query a public resolver (`-Server 1.1.1.1`, `-Server 8.8.8.8`) and compare with the scan. Expect a mix for up to 48 hours | `Export-DomainDns.ps1 -Server 1.1.1.1` |
| 10 | Wait 48 hours | The NS records at the parent zone usually carry a 48-hour TTL. Don't cut over until it has passed | |
| 11 | Clean up | Delete the API key or token you created. Close or hand over the old DNS account if anyone can reach it | |

```powershell
# 2. Scan everything public DNS will tell you
.\Export-DomainDns.ps1 -Domain fabrikam.com -FullScan -Wordlist .\extra-names.txt

# 4. Recreate it at the new provider (preview, then for real). Keys come from environment variables or a prompt
# (set $env:GODADDY_KEY and $env:GODADDY_SECRET first, or let the script prompt; the secret is hidden)
.\Set-DomainDnsRecords.ps1 -Domain fabrikam.com -Provider GoDaddy -FromScan .\DomainDns-FullScan-<date>.csv -WhatIf
.\Set-DomainDnsRecords.ps1 -Domain fabrikam.com -Provider GoDaddy -FromScan .\DomainDns-FullScan-<date>.csv

# 5. Before switching NS: ask the new name servers directly
Resolve-DnsName fabrikam.com -Type MX -Server <new-ns-1>
Resolve-DnsName autodiscover.fabrikam.com -Type CNAME -Server <new-ns-1>

# 9. After switching NS: compare with a public resolver
.\Export-DomainDns.ps1 -Domain fabrikam.com -Server 1.1.1.1
```

## Three ways to do it

| Route | How |
| --- | --- |
| Paid tool | Some DNS providers import a zone file or copy records from another provider in their onboarding wizard. Quick, but check the result record by record: they often miss records they can't see, exactly like a scan |
| Microsoft's own | Azure DNS hosts the zone for pennies a month, has its own alias records for Azure-hosted websites, and can be scripted with Az PowerShell. A good home if the business already has an Azure subscription; one more portal to own if it doesn't |
| By hand | Export a zone file from the old host if anyone can log in; otherwise copy the scan CSV into the new provider's portal one record at a time. Get a second person to check each one |

## Traps

- **Trap: alias records.** Cloud DNS services have "alias" or "flattened" records that let the apex point at a load balancer or CDN. They look like an A record from outside, with an IP address that can change. Copy the IP and the website works today and breaks when the IP moves. Ask the web host what the apex should point at, and use their documented A record or the new provider's own flattening.
- **Trap: the scan can't see everything.** DNS doesn't list a zone's names. Records only answer when asked. Use a wordlist, ask the business, and check old invoices for services that might use a subdomain.
- **Trap: DNSSEC.** If the domain is signed at the old host, switching NS without removing or moving the DS record at the registrar makes the domain fail to resolve. Check before you switch.
- **Trap: NS TTL.** Lowering record TTLs doesn't speed up a name server change. The parent zone's NS TTL (often 48 hours) does what it likes. That's why this happens days before cutover, not hours.
- **Trap: API keys left behind.** A DNS API key can take a business offline. Create it scoped to the one domain where the provider allows, use it, and delete it the same week.
- **Tip:** keep the DNS move and the cutover as separate events with separate rollbacks. If something breaks after the NS switch, you know it's DNS, not Microsoft 365.

## Rollback

Change the name servers at the registrar back to the old list from step 7. The old host still has its records, untouched, so resolution returns to how it was as caches expire (up to 48 hours). Don't delete anything at the old host until cutover is done and stable.

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 1.0.0 | 2026-11-10 | First version, from moving DNS off an inaccessible host two days before a live cutover |
