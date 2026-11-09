# Handover template

**TL;DR:** What I hand the business at the end of a tenant-to-tenant migration: what moved, where things live now, who the admins are, how DNS changed, what's still open and when the old tenant gets switched off. Fill it in at T+7, walk the business owner through it, and keep a copy with the project. Don't put passwords in it.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.0.0 |
| Released | 9 November 2026 |
| Licence | MIT |

Part of the [docs](README.md) for my [M365 tenant-to-tenant migration toolkit](../README.md).

---

## [Company name]: Microsoft 365 migration handover

| | |
| --- | --- |
| Prepared by | [name] |
| Prepared for | [business owner] |
| Date | [date] |
| Cutover date | [date] |
| Support window ends | [date] |
| Source decommission date | [date] |

### 1. Summary

[Two or three sentences in plain English. What moved, when, and whether anything is outstanding.]

### 2. What's moved

| Item | Count | Size | Status | Notes |
| --- | --- | --- | --- | --- |
| User mailboxes | | | | |
| Shared mailboxes | | | | |
| Archives | | | | |
| Distribution lists | | | | |
| Microsoft 365 groups and Teams | | | | |
| SharePoint sites | | | | |
| OneDrive accounts | | | | |
| Mail contacts | | | | |

Attach the final validation report from [09](../09-Validation/) (`Validation-<domain>-<date>.csv`) and summarise its result here: [x pass, y warn, z fail, with each fail explained].

### 3. What didn't move, on purpose

| Item | Why | What happens instead |
| --- | --- | --- |
| Teams chat history | Not moved by the chosen copy option | Agreed in advance. Staff told on [date] |
| Transport rules | Rebuilt by hand and tested | See section 9 |
| [e.g. Leavers' mailboxes] | [reason] | [e.g. Exported and kept until date] |

### 4. Where things are now

| What | Where | Notes |
| --- | --- | --- |
| Email | Outlook, signed in with your normal email address | |
| Shared mailboxes | Appear in Outlook automatically for people with access | |
| Team files | Teams and SharePoint: [list main sites or teams] | |
| Personal files | OneDrive | |
| Admin centre | https://admin.microsoft.com, signed in with an admin account below | |

### 5. Admin accounts and break-glass

Names and roles only. Passwords and recovery details go in your password manager, not in this document.

| Account | Role | Held by | MFA method | Notes |
| --- | --- | --- | --- | --- |
| [admin account] | Global Administrator | | | Day-to-day admin |
| [break-glass account] | Global Administrator | | | Emergency only. Excluded from Conditional Access. Credentials stored at [location] |
| [migration account] | [role] | [name] | | To be removed on [date] |

- [ ] Migration admin accounts removed or disabled on [date]
- [ ] Break-glass sign-in tested on [date]
- [ ] Migration app registrations and consents removed from both tenants on [date]

### 6. DNS before and after

From the `Export-DomainDns.ps1` output ([07](../07-DomainCutover/)) before cutover, and what's set now.

| Domain | Record | Before | After | TTL now |
| --- | --- | --- | --- | --- |
| fabrikam.com | MX | | | |
| fabrikam.com | TXT (SPF) | | | |
| fabrikam.com | CNAME autodiscover | | | |
| fabrikam.com | CNAME selector1._domainkey | | | |
| fabrikam.com | CNAME selector2._domainkey | | | |
| fabrikam.com | TXT _dmarc | | | |
| fabrikam.com | TXT (verification) | | | Can be removed once verified |

- [ ] TTLs raised back to normal on [date]
- [ ] DKIM enabled in the destination on [date]
- [ ] Third-party senders kept in SPF: [list]

### 7. Licences

| Licence | Tenant | Quantity | Assigned | Renewal or end date | Action |
| --- | --- | --- | --- | --- | --- |
| [e.g. Microsoft 365 Business Premium] | Destination | | | | |
| [e.g. Exchange Online Plan 2] | Destination | | | | For shared mailboxes with archives |
| [source licences] | Source | | | | Cancel after decommission |

### 8. Open items

| # | Item | Owner | Due | Status |
| --- | --- | --- | --- | --- |
| 1 | | | | |

### 9. Things to know

- **Transport rules** rebuilt: [list, and whether each is enabled].
- **Apps that send mail** repointed: [list, with who owns each].
- **Security:** [MFA method, Conditional Access policies or security defaults in place].
- **Devices:** [what was done, anything still outstanding].
- **Old file links** to the source no longer work. Staff were told on [date].

### 10. Support window

From [cutover date] to [date], contact [name] at [contact] for anything related to the move. After that, [normal support arrangement].

### 11. Decommissioning the source

| Step | Date | Done |
| --- | --- | --- |
| Source kept read-only for reference | until [date] | |
| Data under hold or retention dealt with: [how] | [date] | |
| Business owner signs off that nothing else is needed from the source | [date] | |
| Source licences cancelled | [date] | |
| Source tenant closed or handed back | [date] | |

### 12. Sign-off

| | Name | Signature | Date |
| --- | --- | --- | --- |
| Business owner | | | |
| Migration lead | | | |
