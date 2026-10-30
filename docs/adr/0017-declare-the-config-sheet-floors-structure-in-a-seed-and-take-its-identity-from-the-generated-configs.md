# Declare the config-sheet floor's structure in a seed and take its identity from the generated configs

The floor seed declares what each config-sheet floor tab looks like and carries no GIDs or column IDs; which tab, Table and column each floor part is comes from the generated configs, so a missing floor tab is recreated at its generated GID and this spreadsheet never migrates. Seeding fixed GIDs was rejected, because the framework serves spreadsheets other than this one and moving this spreadsheet's Value Config to a new GID would break every dropdown that points at it. Every config sync creates a missing floor part from the seed and overwrites a drifted one back to it, failing closed where it can't tell what a person meant, so the floor is enforced rather than fixed by hand, and the generating constructors refuse a floor entry that differs from the seed.

Source: byronbroughten/byro-repo#73
