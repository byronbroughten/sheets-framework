---
status: superseded by sheets-framework#219, which removes the config-sheet floor
---

# Declare the config-sheet floor's structure in a seed and take its identity from the generated configs

The floor seed declares what each config-sheet floor tab looks like and carries no GIDs or column IDs; which tab, Table and column each floor part is comes from the generated configs, so a missing floor tab is recreated at its generated GID and this spreadsheet never migrates. Seeding fixed GIDs was rejected, because the framework serves spreadsheets other than this one and moving this spreadsheet's Value Config to a new GID would break every dropdown that points at it.

## Enforcing the floor

Every config sync creates a missing floor part from the seed and overwrites a drifted one back to it, failing closed where it can't tell what a person meant, so the floor is enforced rather than fixed by hand, and the generating constructors refuse a floor entry that differs from the seed.

## Superseded

[sheets-framework#219](https://github.com/byronbroughten/sheets-framework/issues/219) removes the floor and its seed. No tab, Table or column is the framework's own: each exists because a person made it or an `alterations.ts` entry did, and an ID-prefix guard over every Table takes the place of the floor's identity check.

Source: byronbroughten/byro-repo#73
