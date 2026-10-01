# How Google Sheets Tables work through the Sheets API and Apps Script

Research for byronbroughten/sheets-framework#46. Sources were read on 2026-10-01. Each claim cites its source inline. Labels used below:

- **[doc]**: stated by a Google primary source.
- **[repo-measured]**: not in Google's docs, but measured live against the dev spreadsheet and recorded in this repo.
- **[unverified]**: inference, or something no primary source states. Needs a live probe before anyone relies on it.

Main sources:

- Sheets API Table resource: https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets/sheets (the `Table`, `TableRowsProperties`, `TableColumnProperties`, `ColumnType`, `FilterView`, `BasicFilter` and `ProtectedRange` sections)
- Sheets API requests: https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets/request
- Sheets API Tables guide (last updated 2026-09-03): https://developers.google.com/workspace/sheets/api/guides/tables
- Help, "Use tables in Google Sheets": https://support.google.com/docs/answer/14239833
- Help, "Use table references in Google Sheets": https://support.google.com/docs/answer/15637642
- Apps Script Spreadsheet service index (last updated 2026-04-13): https://developers.google.com/apps-script/reference/spreadsheet
- Apps Script advanced Sheets service: https://developers.google.com/apps-script/advanced/sheets
- Workspace Updates: Tables launch, May 2024 (https://workspaceupdates.googleblog.com/2024/05/tables-in-google-sheets.html); table reference improvements, Nov 2024 (https://workspaceupdates.googleblog.com/2024/11/table-reference-improvements-in-google-sheets.html); API support for tables, weekly recap of 2025-05-02 (https://workspaceupdates.googleblog.com/2025/05/release-notes-05-02-2025.html)

## 1. The Table resource

`Sheet.tables[]` holds "The tables on this sheet" [doc, sheets ref]. Each `Table` has these fields [doc, sheets ref]:

| Field | Type | Doc text |
| --- | --- | --- |
| `tableId` | string | "The id of the table." |
| `name` | string | "The table name. This is unique to all tables in the same spreadsheet." |
| `range` | GridRange | "The table range." |
| `rowsProperties` | TableRowsProperties | "The table rows properties." |
| `columnProperties[]` | TableColumnProperties | "The table column properties." |

- **`TableColumnProperties`** has `columnIndex`, `columnName`, `columnType` and `dataValidationRule`. `columnIndex` is "relative to its position in the table and is not necessarily the same as the column index in the sheet." `dataValidationRule` is "Only set for dropdown column type" [doc, sheets ref].
- **`ColumnType`** values [doc, sheets ref]: `COLUMN_TYPE_UNSPECIFIED`, `DOUBLE`, `CURRENCY`, `PERCENT`, `DATE`, `TIME`, `DATE_TIME`, `TEXT`, `BOOLEAN`, `DROPDOWN`, `FILES_CHIP`, `PEOPLE_CHIP`, `FINANCE_CHIP`, `PLACE_CHIP`, `RATINGS_CHIP`. The UI also has a "None" type for mixed data [doc, help 14239833]. It presumably maps to `COLUMN_TYPE_UNSPECIFIED` or to an absent type [unverified].
- **`TableColumnDataValidationRule`** holds one `condition` (a BooleanCondition), "Valid only if the BooleanCondition.type is ONE_OF_LIST" [doc, sheets ref]. The guide adds: "If a column type is set as dropdown, the dataValidationRule for the column must be set with a ONE_OF_LIST condition. Other column types shouldn't set the dataValidationRule field." Rating and checkbox columns "populate with default values of 0 and FALSE respectively" [doc, guide].
- **`TableRowsProperties`** is colour only: `headerColorStyle`, `firstBandColorStyle`, `secondBandColorStyle` and `footerColorStyle`. The footer has no boolean of its own. Setting `footerColorStyle` adds one. "If updating an existing table without a footer to have a footer, the range will be expanded by 1 row", and removing the footer shrinks the range by 1 row [doc, sheets ref; guide].
- **Header placement.** No field marks the header row. Nothing in the API says the header must be the first row of `range`. The UI and the reference specifiers (`#HEADERS`, `#DATA`, `#TOTALS`) imply this layout: header = first row, footer = last row when present, data = everything between [unverified as an API rule, but consistent with every doc example and with the repo's CONTEXT.md].
- **Other API objects that point at a Table.** `FilterView.tableId` ("The table this filter view is backed by, if any") and `BasicFilter.tableId` take a Table in place of a range. When writing, only one of range / namedRangeId / tableId may be set. `ProtectedRange` also has a `tableId` ("The table this protected range is backed by, if any") [doc, sheets ref]. The guide sums this up: "Other API features that support tables being their backing data include filters, filter views, and protected ranges" [doc, guide].
- **Reading tables.** The repo has measured that a grid fetch returns `tables` only for a data filter that overlaps the Table [repo-measured: `packages/framework/docs/testing.md:41`, `packages/framework/docs/architecture/round-trips.md:12`].

## 2. batchUpdate requests

- **`addTable { table }`**: "Adds a new table to the spreadsheet." `table` is required [doc, request ref]. The guide's example sends `name`, a client-chosen `tableId` ("123"), `range` and `columnProperties` [doc, guide]. A caller can therefore choose the tableId at creation time.
  - The repo found that `columnProperties` sent inside `addTable` write their names three columns to the right of the Table, and the request 500s when those columns don't exist. So the framework sends `addTable` without columns and follows it with an `updateTable` [repo-measured: `packages/framework/src/00_Source/GoogleSheets/GoogleSheetsAPI.ts:413`, `docs/architecture/queued-writes.md:17`]. The guide documents nothing like this, so treat it as a live quirk.
- **`updateTable { table, fields }`**: `fields` is a required FieldMask. "At least one field must be specified. The root table is implied", and `*` covers everything [doc, request ref]. Resizing is done by changing `range` [doc, guide].
  - The repo found that a `columnProperties` list replaces the whole list: an unsent column keeps its header but loses its type [repo-measured: `src/testSupport/fakeSheetsService/tableReplays.ts:66`].
  - The doc text implies that `name` can be updated by mask [doc]. The framework does this for renames (`src/testSupport/fakeSheetsService.test.ts:317`).
- **`deleteTable { tableId }`**: "Removes the table with the given ID from the spreadsheet" [doc, request ref]. The guide says it deletes "the entire table and the contents of the table". To keep the data and drop the formatting, use `DeleteBandingRequest` [doc, guide]. The UI equivalents are "Delete table: Removes the table and its data" and "Revert to unformatted data" [doc, help 14239833].
- **`appendCells { tableId, rows, fields }`**: `tableId` is "The ID of the table to append data to. The data will be only appended to the table body. This field also takes precedence over the sheetId field" [doc, request ref]. The guide: it "appends the values to the first free row and is aware of full rows and footers. If there are no empty rows, this inserts rows to the end of the table and before any footer, if applicable" [doc, guide].
  - The repo found that a Table append inserts rows at the Table's end and shifts the rows below [repo-measured: `src/testSupport/fakeSheetsService/dimensionReplays.ts:13`]. Several appends to the same Table in one batch all target the same first free row, so the framework sends one `appendCells` per Table [repo-measured: `docs/design/one-chokepoint.md:11`].
  - Current use: `GoogleSheetsAPI.ts:285-293` sends `appendCells` with `sheetId`, `tableId` and N empty rows, `fields: "userEnteredValue"`.
- **Inserting and deleting rows and columns.** The guide routes everything through the generic requests: "If you need to add a new row or column within the table, use the InsertRangeRequest or the InsertDimensionRequest. If you need to delete a table row you can use DeleteRangeRequest otherwise you can use DeleteDimensionRequest to delete an entire row from the spreadsheet. Note: If using InsertRangeRequest/DeleteRangeRequest, the range in the request must cover the entire row(s) if deleting a row or entire column(s) if deleting a column" [doc, guide]. No table-specific insert or delete request exists [doc, request ref].

## 3. Several Tables on one sheet

- **Allowed.** The `tables[]` array on `Sheet` [doc, sheets ref] and `name` being unique "to all tables in the same spreadsheet" (not just the sheet) [doc] both allow several Tables per sheet. A Workspace post says Gemini "can now understand and analyze multiple tables within a single tab" (2025, via the blog's Sheets label; seen only in a search snippet) [doc-adjacent].
- **Overlap.** No Google doc states whether Tables may overlap. It's almost certainly refused [unverified; the UI won't let you convert a range that intersects a Table, but nothing written says so].
- **Adjacency.** No doc states whether two Tables may touch side by side or top to bottom with no gap [unverified].
- **Header placement.** See §1. No API rule exists beyond the first row of `range` acting as the header [unverified].
- **Naming rules** (UI). A name can't be TRUE/FALSE, can't look like A1 or R1C1, can't start with a number, can't exceed 255 characters, and can't contain special characters other than `_`. Spaces become `_` in formulas ("Table 1" becomes `Table_1`) [doc, help 14239833]. Whether the API enforces the same rules on `name` is [unverified].

## 4. What dimension changes do to a Table and its neighbours

Google documents almost none of this beyond the guide's routing advice (§2). What the repo has measured:

- An insert strictly inside a Table grows it. An insert at its end grows it only with `inheritFromBefore: true`, which heads the new column `Column <n>`. A delete inside a Table shrinks it. Rows and columns shift Tables along with column types, protected ranges and conditional-format ranges [repo-measured: `docs/testing.md:32`, `src/testSupport/fakeSheetsService/fakeGrid.ts:81`, `:222`].
- Deleting a row above a Table, or inserting a column to its left, moves the Table [repo-measured: `docs/architecture/table-placement.md`].
- An inheriting column insert copies cell format and validation but not the Table column type [repo-measured: `docs/testing.md:32`].

What follows for a **neighbouring** Table, by plain grid geometry (none of it confirmed live):

- `deleteDimension` / `insertDimension` (ROWS) act on whole sheet rows. With two Tables **side by side**, deleting a data row of Table A also deletes the same row index from Table B (a data row, its header, or part of its range). Inserting a row inside A grows B too if the row falls inside B's row span [unverified; inferred from "delete an entire row from the spreadsheet", doc, guide].
- With Tables **stacked**, a row delete or insert in the upper Table shifts the lower one. A table `appendCells` that has to insert rows pushes the lower Table down. That matches the repo's "shifting rows below" measurement [repo-measured for the shift; the effect on a second Table is unverified].
- `deleteRange` / `insertRange` with `shiftDimension: ROWS` over only the Table's columns would, in principle, leave a side-by-side neighbour alone. But the guide says the range "must cover the entire row(s)". That could mean the entire Table row (the useful reading) or the entire sheet row [ambiguous; must be probed live].

## 5. Table references in formulas

Syntax [doc, help 15637642; Workspace Updates Nov 2024]:

| Reference | Meaning |
| --- | --- |
| `Table1[Column 1]` | One column (data) |
| `Table1[[#ALL],[Column 1]]` | One column with header, data and footer |
| `Table1[[#ALL],[Column 1]:[Column 3]]` | Several columns |
| `Table1[#ALL]` | Headers, data and footers |
| `Table1[#HEADERS]` | Headers |
| `Table1[#TOTALS]` | Footer (Google's name for Excel's totals row) |
| `Table1[[#HEADERS],[#DATA]]` | Headers and data |
| `Table1[[#DATA],[#TOTALS]]` | Data and footer |
| `Table1` or `Table1[#DATA]` | Data only |

- **Not supported:** "#This Row currently is not supported" [doc, help 15637642]. So Excel's `[@Col]` / `[#This Row]` current-row form has no documented Google equivalent.
- **Where references are refused:** "These features don't currently support the use of table references when you select a range: Conditional formatting, Charts, Pivot tables" [doc, help 15637642].
- **Other documented uses:** references work in `IMPORTRANGE("url", "DeptSales[Sales Amount]")` and with chip extraction (`Table1[Column 1].[file name]`). To spill a whole Table or column, wrap the reference in `ARRAYFORMULA(Table1[Column1])` [doc, help 15637642].
- **Data changes:** "When you use a name to reference table elements, the references update when you add or remove data from the table" [doc, help 15637642 and 14239833].
- **Column moves, table renames and column renames:** no Google doc states the behaviour. Going in, we knew references survive column moves (repo knowledge, not a doc) [unverified by any doc]. Whether renaming a Table or a column header rewrites dependent formulas is [unverified; needs a live probe]. The framework renames by `updateTable` + `name` and by header writes (`docs/architecture/queued-writes.md:17`), so this matters.
- **`SINGLE`:** absent from Google's function list (https://support.google.com/docs/table/25273, searched 2026-10-01), and no help page documents it. The repo uses `=2+SINGLE(test[Number])` (`src/02_SpreadsheetRaw/SpreadsheetRaw.test.ts:2504`, `src/04_SpreadsheetNamed/SpreadsheetNamed.test.ts:817`), and CONTEXT.md describes it as the wrapper used "when one cell is wanted". It most likely behaves like Excel's `SINGLE` / `@` implicit intersection: it picks the cell of the referenced column in the formula's own row, which is the practical stand-in for the unsupported `#This Row` [unverified; undocumented by Google, working by observation only].
- **Formula help in the editor:** formula suggestions flag formulas whose ranges overlap a Table, and formula corrections offer "a proper table reference" [doc, Workspace Updates Nov 2024].

## 6. Apps Script visibility

- **SpreadsheetApp has no Table class.** The Spreadsheet service's class list (last updated 2026-04-13) includes `PivotTable`, `DataSourceTable` (Connected Sheets/BigQuery), `EmbeddedTableChartBuilder`, `NamedRange`, `Protection` and so on, but no Table and no Table methods on `Sheet` or `Range` [doc, Apps Script spreadsheet index]. A `Table` class exists only in the Document and Slides services [doc, same page]. So `SpreadsheetApp` can't list, create, resize or read the type of a Sheets Table. Plain cell writes through `Range` still land inside a Table's range like any others [unverified detail: whether a `Range.insertRows` behaves like `insertDimension` with respect to Table growth].
- **The advanced Sheets service sees everything the API does.** It "uses the same objects, methods, and parameters as the public API" [doc, advanced/sheets]. `Sheets.Spreadsheets.get(...)` returns `sheets[].tables[]`, and `Sheets.Spreadsheets.batchUpdate({requests}, id)` accepts `addTable` / `updateTable` / `deleteTable` / `appendCells.tableId` [doc, by the same-objects statement]. The repo's types already use `GoogleAppsScript.Sheets.Schema.AppendCellsRequest` (`src/testSupport/fakeSheetsService/dimensionReplays.ts:16`).
- **Timeline:** Tables launched in the UI in May–June 2024 [doc, Workspace Updates May 2024]. API support arrived around 2025-05-02: "Following the improvements made to tables in Google Sheets in March and April, we're excited to introduce API support for tables," described as "basic actions to create and modify tables" [doc, Workspace Updates 2025-05-02]. The March and April 2025 posts added table formatting (footer toggle, gridlines, condensed view) and group-by aggregation [doc, recaps of 2025-03-07 and 2025-04-04].

## 7. What this means for a table-centric abstraction

**Constraints**

- **Identify a Table by `tableId`, look it up by `name`.** `tableId` is the only handle that `deleteTable`, `appendCells`, filters, filter views and protected ranges accept [doc]. A user can rename a Table in the UI ("Rename table") [doc, help 14239833], and the name is spreadsheet-unique [doc]. So `spreadsheet.table(name)` is a good public lookup, but the code should resolve the name to the live `tableId` on each fetch and key every write and every bit of cached state on that id. `addTable` lets the caller choose the id [doc, guide], and the framework already sets `tableId = name` at creation (`GoogleSheetsAPI.ts:419`). After a UI rename, that id no longer matches the name, so code must never assume id == name.
- **Row deletes on a multi-table sheet are the main hazard.** `deleteDimension` removes whole sheet rows [doc]. The current blank-row and delete path assumes one Table per sheet (`docs/architecture/table-placement.md`). Two options:
  - Require stacked-only layouts and recompute each Table's range after every delete in a batch.
  - Probe whether `deleteRange` with `shiftDimension: ROWS`, bounded to one Table's columns, is accepted and leaves side-by-side neighbours alone (§4, ambiguous doc).

  Until that probe runs, "one Table per sheet, or Tables only stacked" is the safe rule.
- **Appends shift what lies below.** A table-aware append that inserts rows pushes a lower Table down, so any cached range for that Table goes stale within the same flush. Range bookkeeping should be per sheet, not per Table.
- **Formulas can't use `#This Row`, and `SINGLE` is undocumented.** A generator that emits `SINGLE(T[Col])` depends on undocumented behaviour. Pin it with a live test. Table references are also unusable in conditional formats, charts and pivots [doc], so those features still need A1 ranges derived from `tables.range`.
- **Apps Script has no Table class.** All Table work has to go through the advanced service or REST, which the framework already does.

**Opportunities**

- **A generated `tableConfigs` maps cleanly onto the API.** Name, column names, `columnType` and dropdown `ONE_OF_LIST` values are all first-class `Table` fields [doc]. A config sync can diff the live `columnProperties` against the generated ones and send one full-list `updateTable` per Table (full list, because a partial list drops types [repo-measured]). Store the `tableId` in the generated config so a UI rename shows up as a name drift rather than a missing Table.
- **One `appendCells` per Table, with the ordering already measured,** gives cheap, footer-aware inserts [doc + repo-measured].
- **Filters, filter views and protections can bind to `tableId`** [doc]. If the framework wants edit locks, they could follow the Table as it grows rather than being re-ranged by hand. Whether a table-backed protected range actually tracks growth is [unverified].
- **The footer (`#TOTALS`)** is a supported, typed place for aggregate formulas (`SUM(T[Col])`) that the API can toggle via `footerColorStyle` [doc]. Toggling it changes `range` by one row, which any range math must account for.

## Open questions to probe live (dev spreadsheet)

1. Does `addTable` refuse a range that overlaps another Table, or one that touches it?
2. With two Tables side by side, does `deleteRange` with `shiftDimension: ROWS`, bounded to Table A's columns, succeed and leave Table B intact?
3. Does renaming a Table (`updateTable` with `name`) or a header cell rewrite dependent `T[Col]` formulas? Does moving a column?
4. What is `SINGLE`'s exact semantics, and does it survive a round trip through `userEnteredValue`?
5. Does the API apply the UI naming rules to `name`?
