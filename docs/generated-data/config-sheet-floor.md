# The config-sheet floor

Part of [generated data](../generated-data.md). The seed declares structure and the generated files hold identity; every config sync repairs the floor, and a regeneration refuses a changed floor identity or an entry that differs from the seed.

## The seed declares structure, the generated files hold identity

**The config-sheet floor: the seed declares structure, the generated files hold identity.** The four config-describing sheets — `sheetConfig`, `columnConfig`, `spreadsheetConfig`, and `valueConfig` — form a floor underneath the rest of the generated data, because config operators need a fixed way to find those sheets and their own columns (including Spreadsheet Config's GID and guaranteed headers) before they can generate anything else. The floor seed (`01_SpreadsheetSchema/configSheetFloorSeed.ts`) declares what each floor tab looks like and carries no GIDs or column IDs; the `sheetConfigs`/`columnConfigs` entries for those sheets hold which tab and column each one is. A missing floor part is created from the seed at its generated identity, and a drifted one is overwritten back to the seed. `spreadsheetConfig` the generated layout record is not part of that floor.

## What every config sync repairs

The live Spreadsheet Config, Sheet Config and Column Config tabs each carry one whole-sheet edit warning with editable ranges where an edit sticks, and a floor column whose type differs from the seed is set back and reported, both applied at the start of every config sync (and by the `ensureConfigSheetFloor` chore). Tab titles, Table names, headers, column IDs and column-group headings found at a floor GID are overwritten back to the seed on that same run; extra columns people added are left alone. A deleted floor column the sync or an endpoint refills by itself (Sheet Config's Sheet title, Column Config's Sheet title and Header, and Spreadsheet Config's endpoint feedback columns) is recreated at its Table's end on that run; any other missing floor column, or a Table menu space that is no longer the Table's first column, stops the sync with a message naming it and flushes nothing. A column type is restored only on a Table with no dropdown or other validated column; otherwise the run stops and names those columns, because restoring a type resets every dropdown's style and colours on that Table.

## A deleted floor tab is recreated

A deleted Spreadsheet Config, Sheet Config or Column Config tab is created at its generated GID with its seed title and a Table of its seeded columns (Spreadsheet Config's endpoint feedback columns included), placed by the generated layout with one data row, empty except that a created Spreadsheet Config's layout values are seeded in the same flush from that generated layout, in base 1, and never restored afterwards; on that run the new tab's column IDs, headings and edit warning follow from the refetch, and its editable ranges are carved on the next sync. A deleted Value Config is created at its generated GID with its seed title and a Table holding only the seed's example column (`Example value`) over two data rows; in the same flush its two sample members are seeded into those rows and its column ID is minted with Value Config's generated ID prefix (`vcf`), since it has no floor columns to carry the prefix back. That column is a seeded value, written only when the Table is created and never restored, so a deleted example stays deleted; Value Config gets no edit warning.

## A regeneration refuses a changed floor identity

A floor tab's GID and ID prefix and a floor column's column ID are its identity, so `gen:configs` compares each with the previous generated files before writing the four and refuses, naming what changed and what it was, when one differs; a floor tab or column absent from the previous files is skipped, and a non-floor sheet's changed ID prefix is still only reported.

## A regeneration refuses a floor entry that differs from the seed

After that identity check, emit refuses a floor entry that differs from the floor seed (`assertFloorMatchesSeed`, which runs at regen rather than inside the generating constructors, since it resolves column IDs through the previous generated files): it names a floor tab with no entry (its Let api access unticked), and names the sheet, the column and both values for a seeded floor column that is missing, or whose header, `valueName` (the one its column type implies) or `emptyValueAllowed` differs. The comparison runs one way, so live columns the seed doesn't declare and Custom default value are left alone. Changing the ID delimiter therefore fails the regen, since the floor's column IDs keep the delimiter they were created with.

## Never hand-edit a floor entry

The floor entries in the generated files must still not be hand-edited: fix the live tab, or the seed, and let the floor and the regen carry it.
