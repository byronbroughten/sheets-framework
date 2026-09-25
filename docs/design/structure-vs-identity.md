# Structure is declared in code; identity is recorded from the sheet

Design reasoning. The one-line principle lives in [`docs/design.md`](../design.md); this file holds the argument and its instances.

## The argument

What a thing looks like — headers, floor columns, ID shape — is declared in code. Which instance it is — GID, column ID, prefix — is read from the sheet.

## Instances

The floor seed says what the config-sheet floor looks like and carries no GIDs or column IDs; which tab and column it is comes from the generated configs, and a missing floor tab is recreated at its generated GID. Seeding fixed GIDs was rejected: the framework serves spreadsheets other than this one, and moving this spreadsheet's Value Config to a new GID would break every dropdown that points at it (#73). An ID prefix is identity, so it is read from the sheet's column IDs rather than declared in code or in the generated-config cache; using the cache as an input to its own regeneration would lose prefixes on a reset and let two branches assign the same one (#87).
