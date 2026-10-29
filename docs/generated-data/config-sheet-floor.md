# The config-sheet floor

Part of [generated data](../generated-data.md). The seed declares structure and the generated files hold identity; every config sync repairs the floor, and a regeneration refuses a changed floor identity or an entry that differs from the seed.

## The seed declares structure, the generated files hold identity

**The config-sheet floor: the seed declares structure, the generated files hold identity.** The three config-describing sheets — `tableConfig`, `columnConfig` and `spreadsheetConfig` — form a floor underneath the rest of the generated data, because config operators need a fixed way to find those sheets and their own columns (including Spreadsheet Config's GID and guaranteed headers) before they can generate anything else. The floor seed (`01_SpreadsheetSchema/configReaders/configSheetFloorSeed.ts`) declares what each floor tab looks like and carries no GIDs, Table IDs or column IDs; the `tableConfigs`/`columnConfigs` entries for those Tables hold which tab, Table and column each one is. A missing floor part is created from the seed at the identity the generated files hold, and a drifted one is overwritten back to the seed.

## What every config sync repairs

The live Spreadsheet Config, Table Config and Column Config tabs each carry one whole-sheet edit warning with editable ranges where an edit sticks, and a floor column whose type differs from the seed is set back and reported, both applied at the start of every config sync (and by the `ensureConfigSheetFloor` chore). Tab titles, Table names, headers, column IDs and group headings found at a floor GID are overwritten back to the seed on that same run; extra columns people added are left alone. A floor tab holds only its floor Table, the one at its generated Table ID: a sync that finds another Table on it stops before writing anything, naming that Table. The three floor Tables' Table Config rows are self-describing: their Let api access is enforced TRUE, so a sync that finds one unticked ticks it again and reports it. A deleted floor column the sync or an endpoint refills by itself (Table Config's Table name and Sheet title, Column Config's Table name and Header, and Spreadsheet Config's endpoint feedback columns) is recreated at its Table's end on that run; any other missing floor column, or a Table menu space that is no longer the Table's first column, stops the sync with a message naming it and flushes nothing. A column type is restored only on a Table with no dropdown or other validated column; otherwise the run stops and names those columns, because restoring a type resets every dropdown's style and colours on that Table.

## A deleted floor tab is recreated

A deleted Spreadsheet Config, Table Config or Column Config tab is created at its generated GID with its seed title and a Table, at its generated Table ID, of its seeded columns (Spreadsheet Config's endpoint feedback columns included), placed by the Table layout (the `tableLayout` constant) with one empty data row; on that run the new tab's column IDs, headings and edit warning follow from the refetch, and its editable ranges are carved on the next sync. Value Config is not on the floor (sheets-framework#148): a deleted one is not recreated, and an existing one is an ordinary Table, identified from the sheet like any other, its ID prefix included.

## A regeneration refuses a changed floor identity

A floor tab's GID, Table ID and ID prefix and a floor column's column ID are its identity, so `gen:configs` compares each with the previous generated files before writing them and refuses, naming what changed and what it was, when one differs; a floor tab or column absent from the previous files is skipped, and a non-floor sheet's changed ID prefix is still only reported.

## A regeneration refuses a floor entry that differs from the seed

After that identity check, emit refuses a floor entry that differs from the floor seed (`assertFloorMatchesSeed`, which runs at regen rather than inside the generating constructors, since it resolves column IDs through the previous generated files): it names a floor tab with no entry (its Let api access unticked), and names the sheet, the column and both values for a seeded floor column that is missing, or whose header, `valueName` (the one its column type implies) or `emptyValueAllowed` differs. The comparison runs one way, so live columns the seed doesn't declare are left alone. The ID delimiter lives in the `tableLayout` constant, not on a sheet ([the rule](../../src/01_SpreadsheetSchema/AGENTS.md)).
