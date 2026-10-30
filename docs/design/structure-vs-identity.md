# Structure is declared in code; identity is recorded from the sheet

Design reasoning. The one-line principle lives in [`docs/design.md`](../design.md); this file holds the argument and its instances.

## The argument

What a thing looks like — headers, floor columns, head rows, ID shape — is declared in code. Which instance it is — GID, `tableId`, column ID, prefix — is read from the sheet.

## Instances

[ADR 0017](../adr/0017-declare-the-config-sheet-floors-structure-in-a-seed-and-take-its-identity-from-the-generated-configs.md) splits the config-sheet floor along this line, and rejects seeding fixed GIDs because the framework serves other spreadsheets than this one. An ID prefix is identity, so it is read from the sheet's column IDs rather than declared in code or in the generated-config cache; using the cache as an input to its own regeneration would lose prefixes on a reset and let two branches assign the same one ([ADR 0012](../adr/0012-assign-each-sheets-id-prefix-once-and-sample-it-from-its-column-ids.md) samples the prefix from the column ID row and rejects the cache as its source). Every Table's head (which head rows sit above its header, and at what offsets) and its ID shape (the delimiter, the ID and Name headers) are structure, so they are declared in one code constant, `tableLayout`, rather than read from Spreadsheet Config's cells (#6; sheets-framework#50). A `tableId` is identity even when the framework mints it: a Table the framework creates gets a random `tableId` of its own, unrelated to its name, and from then on that ID is read from the sheet like any other (sheets-framework#51, sheets-framework#59).
