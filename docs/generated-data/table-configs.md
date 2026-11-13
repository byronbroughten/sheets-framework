# `tableConfigs` and Table Config: the Table list

Part of [generated data](../generated-data.md). Table Config lists every Google Table in the spreadsheet, and `tableConfigs` holds the ones with API access; `hasIdColumn`, `hasNameColumn` and `idPrefix` are sampled from the described Table. Beside it, `generated/tabConfigs.ts` holds `tabConfigs`, the tab list, and `generated/retiredIdPrefixes.ts` holds `retiredIdPrefixes`.

## The Table Config sheet

**Table Config has one row per Google Table in the spreadsheet: `Table ID | Table name | Sheet title | Let api access`.** `TableConfigsOperator` (`05_Operators`) owns it. Table ID is the row's identity. Table name and Sheet title are readable identity for a human. Let api access is the only declared column, and nothing reads it. The regen neither reads nor writes this sheet ([config sheets](./config-sync.md)); API access is a Table's presence in the previous `tableConfigs`, until a `letApiAccess` alteration changes it ([granting and revoking](#granting-and-revoking-api-access)). A tab with two Tables gets two rows, and a tab with no Table gets none.

## How `tableConfigs` is generated

**`tableConfigs`** (`generated/tableConfigs.ts`) holds one entry per Table with API access, keyed by the live Table name through `titleToName`. A Table without it gets no entry and no Column Config rows, and two Tables whose names give the same key fail the run, naming both. Each entry holds the Table's live `tableId` and name, its sheet GID, its `idPrefix`, `hasIdColumn` and `hasNameColumn`. `sheetGid` is recorded, but a regen never refuses a Table because it changed: a Table moved to another tab is the same Table. It records no position: a run finds the Table by its `tableId` in the sheet's header zone ([table placement](../architecture/table-placement.md)), and the run refuses a Table with API access whose header row is outside it.

## Granting and revoking API access

**`letApiAccess: true` in `alterations.ts` grants API access, and `letApiAccess: false` revokes it.** A grant's `--send` writes the Table's column IDs on its column ID row and takes it into `tableConfigs`; a dry run plans it and writes nothing. A grant takes its own run: any other entry naming that Table is refused before reading, while other Tables' entries share the run. A grant on a Table whose header row is outside the header zone stops the run, naming the Table and saying to move it. Any other entry naming a Table without API access is refused before reading, saying to grant `letApiAccess` first. A revoke drops the Table from the configs and leaves it on the sheet.

**`hasIdColumn` is *sampled*** at emit time from whether the Table's header row contains the ID header (`tableLayout`'s `ID`). **`hasNameColumn` is sampled** the same way, from the `Name` header; a Table with both is one `TableNamed.rowIdByName` can search. **`idPrefix` is *sampled*** from the column ID row, generated from the Table name when that row has no column IDs yet, and checked against the previous `tableConfigs` entry with the same `tableId`. A difference is reported, not failed. ID prefixes must be unique across entries, and `makeTableConfigs` and the emit share that check.

## `tabConfigs`: the tab list

**`tabConfigs`** holds one entry per tab, keyed by its title through `titleToName`: `sheetGid`, `title`, `color`, `frozenRows`, `frozenColumns` and `hidden`. `color` is `#rrggbb`, a theme colour's name such as `ACCENT1`, or `null` for none. Two tab titles that give the same tab name fail the run, naming both. It records no order: order is never in the baseline, for tabs or columns. The fields come from the sheet-properties read every run already makes, so the list costs no extra round trip.

## `retiredIdPrefixes`: prefixes no new Table may take

**`retiredIdPrefixes`** keeps an old row ID from ever pointing at a new Table. Each run carries the previous list forward, adds the `idPrefix` of each previous `tableConfigs` entry that has lost API access or no longer exists, and drops any prefix a Table with API access holds now, so a Table granted again keeps its own. When a Table with no column IDs gets a new prefix, `idPrefixes.fromTitle` treats every retired prefix as taken, as it does the prefix of each last-run Table gone from this one.

## Column references

`TableConfigsOperator.parseColumnReference` resolves a structured reference such as `Rents[Tenant]` to `(tableId, columnId)` against the live names of the Tables with API access and headers. No dropdown validator is built on it yet.
