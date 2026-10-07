# `tableConfigs` and Table Config: the Table list

Part of [generated data](../generated-data.md). Table Config lists every Google Table in the spreadsheet, and `tableConfigs` holds the ones ticked Let api access; `hasIdColumn`, `hasNameColumn` and `idPrefix` are sampled from the described Table.

## The Table Config sheet

**Table Config has one row per Google Table in the spreadsheet: `Table ID | Table name | Sheet title | Let api access`.** `TableConfigOperator` (`05_Operators`) owns it. Table ID is the row's identity. Table name and Sheet title are readable identity, which the correction pass rewrites from the live Table and its tab. Let api access is the only declared column: a human's tick, except on a **self-describing row**, where the correction pass writes the floor seed. The sync appends a new Table's row unticked and deletes the row of a Table that no longer exists. A tab with two Tables gets two rows, and a tab with no Table gets none.

## How `tableConfigs` is generated

**`tableConfigs`** (`generated/tableConfigs.ts`) holds one entry per ticked Table, keyed by the live Table name through `titleToName`. An unticked Table gets no entry and no Column Config rows, and two Tables whose names give the same key fail the run, naming both. Each entry holds the Table's live `tableId` and name, its sheet GID, its `idPrefix`, `hasIdColumn` and `hasNameColumn`, and its `headerRowIndex` and `startColIndex` as recorded at generation.

**`hasIdColumn` is *sampled*** at emit time from whether the Table's header row contains the ID header (`sheetLayout`'s `ID`). **`hasNameColumn` is sampled** the same way, from the `Name` header; a Table with both is one `TableNamed.rowIdByName` can search. **`idPrefix` is *sampled*** from the column ID row, generated from the Table name when that row has no column IDs yet, and checked against the previous `tableConfigs` entry with the same `tableId`. A difference is reported, not failed. ID prefixes must be unique across entries, and `makeTableConfigs` and the emit share that check.

## `sheetConfigs`, until it is removed

**`sheetConfigs`** (`generated/sheetConfigs.ts`) is still emitted beside it: one entry per sheet that holds a ticked Table, keyed by the sheet title, with that Table's GID, `idPrefix`, `hasIdColumn` and `hasNameColumn`. Column Config and the floor still read it, and it goes once they read `tableConfigs`.

## Column references

`TableConfigOperator.parseColumnReference` resolves a structured reference such as `Rents[Tenant]` to `(tableId, columnId)` against the ticked Tables' live names and headers. No dropdown validator is built on it yet.
