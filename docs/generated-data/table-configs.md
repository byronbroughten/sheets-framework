# `tableConfigs` and Table Config: the Table list

Part of [generated data](../generated-data.md). Table Config lists every Google Table in the spreadsheet, and `tableConfigs` holds the ones ticked Let api access; `hasIdColumn`, `hasNameColumn` and `idPrefix` are sampled from the described Table. The same file, `generated/tableConfigs.ts`, holds `tabConfigs`, the tab list, and `retiredIdPrefixes`.

## The Table Config sheet

**Table Config has one row per Google Table in the spreadsheet: `Table ID | Table name | Sheet title | Let api access`.** `TableConfigOperator` (`05_Operators`) owns it. Table ID is the row's identity. Table name and Sheet title are readable identity, which the correction pass rewrites from the live Table and its tab. Let api access is the only declared column: a human's tick, except on a **self-describing row**, where the correction pass writes the floor seed. The sync appends a new Table's row unticked and deletes the row of a Table that no longer exists. A tab with two Tables gets two rows, and a tab with no Table gets none.

## How `tableConfigs` is generated

**`tableConfigs`** (`generated/tableConfigs.ts`) holds one entry per ticked Table, keyed by the live Table name through `titleToName`. An unticked Table gets no entry and no Column Config rows, and two Tables whose names give the same key fail the run, naming both. Each entry holds the Table's live `tableId` and name, its sheet GID, its `idPrefix`, `hasIdColumn` and `hasNameColumn`. `sheetGid` is a trait, never a stop: a Table moved to another tab is the same Table. It records no position: a run finds the Table by its `tableId` in the sheet's header zone ([table placement](../architecture/table-placement.md)), and the run refuses a ticked Table whose header row is outside it.

**`hasIdColumn` is *sampled*** at emit time from whether the Table's header row contains the ID header (`tableLayout`'s `ID`). **`hasNameColumn` is sampled** the same way, from the `Name` header; a Table with both is one `TableNamed.rowIdByName` can search. **`idPrefix` is *sampled*** from the column ID row, generated from the Table name when that row has no column IDs yet, and checked against the previous `tableConfigs` entry with the same `tableId`. A difference is reported, not failed. ID prefixes must be unique across entries, and `makeTableConfigs` and the emit share that check.

## `tabConfigs`: the tab list

**`tabConfigs`** holds one entry per tab, keyed by its title through `titleToName`: `sheetGid`, `title`, `color`, `frozenRows`, `frozenColumns` and `hidden`. `color` is `#rrggbb`, a theme colour's name such as `ACCENT1`, or `null` for none. Two tab titles that give the same tab name fail the run, naming both. It records no order: order is never in the baseline, for tabs or columns. The fields come from the sheet-properties read every run already makes, so the list costs no extra round trip.

## `retiredIdPrefixes`: prefixes no new Table may take

**`retiredIdPrefixes`** keeps an old row ID from ever pointing at a new Table. Each run carries the previous list forward, adds the `idPrefix` of each previous `tableConfigs` entry that is no longer ticked or no longer exists, and drops any prefix a ticked Table holds now, so a Table ticked again keeps its own. When a Table with no column IDs gets a new prefix, `idPrefixes.fromTitle` treats every retired prefix as taken, as it does the prefix of each last-run Table gone from this one.

## Column references

`TableConfigOperator.parseColumnReference` resolves a structured reference such as `Rents[Tenant]` to `(tableId, columnId)` against the ticked Tables' live names and headers. No dropdown validator is built on it yet.
