# 02b User mapping

**TL;DR:** One file says who is who between the two tenants, and every later step depends on it. `New-UserMapping.ps1` reads both tenants, matches every mailbox, list and contact, and writes the mapping CSV, a ShareGate `.sgum`, and a short list of the matches a human should check. It changes nothing in either tenant.

| | |
|---|---|
| Author | Mezba Uddin (Microsoft MVP) |
| Version | 1.1.0 |
| Runs against | Both tenants, read-only |
| Needs | ExchangeOnlineManagement 3.x, View-Only Recipients (or higher) in both tenants |

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
# After 02 has built the destination mailboxes
.\New-UserMapping.ps1 -SourceAdmin admin@fabrikam.com -DestinationAdmin admin@contoso.com

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

## Lessons learned

- **Run it twice.** Once after 02, and again after [07](../07-DomainCutover/). Sign-in names change at cutover, and the `.sgum` maps by sign-in name. An incremental with the old file maps nobody.
- **Test it against reality before trusting it.** My first version agreed with our hand-built mapping on 272 of 278. The other 6 found two real flaws: it matched people to their own guest accounts, and it matched a shared mailbox to another company's mailbox of the same name in the destination. Hence the guest rule, the build plan and `-DestinationDomain`.
- **Display name is a last resort.** Two "Pat Jones" is more common than you'd think. The script refuses a name match that isn't unique and flags every name match it does make.
- **Leavers belong in the overrides file**, not left unmatched. "Unmatched" should mean "I haven't decided yet", so the review list stays useful.
- **Keep the output out of source control.** It's a list of every person and address in two companies.

Not sponsored by ShareGate. The `.sgum` is there because that's the tool I used on the live migration.
