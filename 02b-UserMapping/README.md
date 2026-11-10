# 02b User mapping

**TL;DR:** One file says who is who between the two tenants, and every later step depends on it. `New-UserMapping.ps1` reads both tenants, matches every mailbox, list and contact, and writes the mapping CSV, a ShareGate `.sgum`, and a short list of the matches a human should check. It changes nothing in either tenant. Run it after 02, and again after cutover. Examples use fabrikam.com as the source and contoso.com as the destination.

| | |
| --- | --- |
| Author | Mezba Uddin, Microsoft MVP |
| Version | 1.1.0 |
| Released | 9 November 2026 |
| Licence | MIT |

Part 2b of my [M365 tenant-to-tenant migration toolkit](../README.md). Run it once [New-TargetMailboxes](../02-TargetMailboxes/) has built the destination.

## Requirements

- `ExchangeOnlineManagement` 3.x
- View-Only Recipients (or higher) in both tenants. It only reads

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

## Why you need it

The mapping (`Source,Destination`: old address to new sign-in) drives:

- [05 Mail config](../05-MailConfig/): who gets Full Access and Send As on which shared mailbox, who's in which list, where forwarding goes
- [06 Groups, Teams and sites](../06-GroupsTeamsSites/): owners and members
- [07 Domain cutover](../07-DomainCutover/): which old addresses and X500s land on which new mailbox
- [09 Validation](../09-Validation/): what to check against what
- Your copy tool's user mapping, so file authors, "modified by" and permissions point at real people

## What happens without it

You build it by hand from two exports. It works until someone has a different alias in the new tenant, two people share a display name, or a shared mailbox was renamed. Then:

- Permissions go to the wrong person, or to nobody, and a shared mailbox arrives that nobody can open.
- The copy tool can't resolve authors, so every file shows as created by the migration account.
- After cutover, sign-in names change and the hand-built ShareGate file quietly stops matching anyone on the incrementals.

## How it matches

In this order, first hit wins:

| # | Method | Notes |
|---|---|---|
| 1 | Overrides file | Your decisions. A blank Destination leaves that source out (leavers, test accounts) |
| 2 | Build plan | Part 02's `TargetBuildPlan.csv` (`-PlanFile`). The best source for mailboxes you staged under a different alias |
| 3 | Toolkit tag | Lists and contacts created by 05 and 06 carry the source address in `CustomAttribute15` |
| 4 | Old address | The destination already holds the source address as an alias (staged mailboxes, after cutover) |
| 5 | Alias | Tries `-AliasPrefix` + alias first. Only if exactly one destination object has it |
| 6 | Address local part | `alex@fabrikam.com` to `alex@contoso.com`, same prefix rule, only if unique |
| 7 | Display name | Only if unique, and always flagged for review |

Rules on top:

- **A guest is never a match.** Guests in the destination often carry the person's old address, so step 4 would happily pick them.
- **A mailbox only maps to a mailbox.**
- **`-DestinationDomain` scopes steps 5 to 7** to the domains this migration owns. Without it, in a destination with several domains, every alias and local part guess is flagged for review.
- Two sources mapped to the same destination are flagged.

## Run it

```powershell
# After 02 has built the destination mailboxes. The source still owns fabrikam.com
.\New-UserMapping.ps1 -SourceAdmin admin@fabrikam.com -DestinationAdmin admin@contoso.onmicrosoft.com

# After cutover. fabrikam.com now lives in the destination, so sign in to the source with its onmicrosoft.com name
.\New-UserMapping.ps1 -SourceAdmin admin@fabrikam.onmicrosoft.com -DestinationAdmin admin@contoso.onmicrosoft.com

# Already have a hand-built one? See where they disagree
.\New-UserMapping.ps1 -CompareWith .\UserMapping-old.csv

# Your decisions for the ones it couldn't work out
.\New-UserMapping.ps1 -OverridesFile .\MappingOverrides.csv

# Destination already hosts other companies, staged objects prefixed 'fab-'
.\New-UserMapping.ps1 -PlanFile .\TargetBuildPlan.csv -DestinationDomain contoso.com, fabstage.com -AliasPrefix fab-
```

`MappingOverrides.csv`:

```text
Source,Destination
pat@fabrikam.com,pjones@contoso.com
leaver@fabrikam.com,
```

Output:

| File | Use |
|---|---|
| `UserMapping.csv` | The mapping for 05, 06, 07, 09. Matched rows only |
| `UserMapping.sgum` | Import into ShareGate's user and group mapping. Mailbox users, by sign-in name |
| `UserMapping-Review.csv` | Every source recipient, how it matched, and the ones to check first |

Fix what's in the review list with the overrides file and run it again until the review list is empty or you're happy with it.

## Three ways to do it

| Route | How |
| --- | --- |
| Paid tool | ShareGate, and most copy tools, have a user mapping screen that matches by name or address and takes a CSV or its own file (`.sgum` for ShareGate). Good for content authors. It doesn't know about your staged aliases, guests or another company's `info@`, and its file stops matching after cutover unless you rebuild it |
| Microsoft's own | There's no mapping tool between two tenants. For the native route, the mapping is the mail users you create in the destination, so their addresses are the mapping. Export recipients from both Exchange admin centres and match them in a spreadsheet |
| By hand | Two exports, `Get-Recipient` in each tenant, matched with `VLOOKUP` on alias or address, then fixed row by row. Fine for twenty people. Check every row where two people share a name |

## Tips and traps

- **Tip:** sign in to the source with its onmicrosoft.com admin after cutover. `admin@fabrikam.com` no longer exists in the source once 07 has moved the domain.
- **Tip:** keep `MappingOverrides.csv` with the project files. It's the record of every decision a human made, and the next run needs it.
- **Tip:** run it again after [New-TargetResourceMailboxes](../02-TargetMailboxes/) so rooms and equipment are in the mapping. They carry the toolkit tag, so they match on step 3.
- **Trap:** guests. A destination guest with the person's old address looks like a perfect match. The script never picks a guest, but a hand-built file often does.
- **Trap:** another company's `info@`. In a destination that hosts several companies, alias guessing finds the wrong one. Pass `-DestinationDomain`.
- **Trap:** an old `.sgum` after cutover. It maps by sign-in name, the names changed, and an incremental with it maps nobody: every file shows as created by the migration account.

## Lessons learned

- **Run it twice.** Once after 02, and again after [07](../07-DomainCutover/). Sign-in names change at cutover, and the `.sgum` maps by sign-in name. An incremental with the old file maps nobody.
- **Test it against reality before trusting it.** My first version agreed with our hand-built mapping on 272 of 278. The other 6 found two real flaws: it matched people to their own guest accounts, and it matched a shared mailbox to another company's mailbox of the same name in the destination. Hence the guest rule, the build plan and `-DestinationDomain`.
- **Display name is a last resort.** Two "Pat Jones" is more common than you'd think. The script refuses a name match that isn't unique and flags every name match it does make.
- **Leavers belong in the overrides file**, not left unmatched. "Unmatched" should mean "I haven't decided yet", so the review list stays useful.
- **Keep the output out of source control.** It's a list of every person and address in two companies.

Not sponsored by ShareGate. The `.sgum` is there because that's the tool I used on the live migration.

## Changelog

| Version | Date | Change |
| --- | --- | --- |
| 1.1.0 | 2026-11-09 | Never maps to a guest. `-DestinationDomain` scopes alias, local part and name guesses to the domains this migration owns. `-PlanFile` uses part 02's build plan first. `-AliasPrefix` |
| 1.0.0 | 2026-11-09 | First release |
