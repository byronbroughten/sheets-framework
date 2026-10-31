# Which Sheets API request sets each declared property on a Google Table, and with what side effects

Research for byronbroughten/sheets-framework#179 (part of #178), on 2026-10-31. The question: for each property a sparse declarations file might carry, which Sheets API v4 `batchUpdate` request sets it on a Google Table or its column, what that request does to neighbouring state, what the API refuses, which framework code already emits it, whether the test fake models it, and whether a preview could tell that it touches typed (non-formula) data.

How much to trust each claim:

- **[doc]** comes from a Google primary source, cited inline. Short names are used for these:
  - Tables guide: <https://developers.google.com/workspace/sheets/api/guides/tables> (last updated 2026-09-03).
  - Sheets ref: <https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets/sheets> (the `Table`, `TableColumnProperties`, `ProtectedRange` and `ConditionalFormatRule` schemas).
  - Request ref: <https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets/request>.
  - Cells ref: <https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets/cells>.
  - Other ref: <https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets/other> (`BooleanCondition` / `ConditionType`).
  - Response ref: <https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets/response>.
  - Field masks guide: <https://developers.google.com/workspace/sheets/api/guides/field-masks>.
  - Help: tables: <https://support.google.com/docs/answer/14239833>.
  - Help: table refs: <https://support.google.com/docs/answer/15637642>.
  - WU 2024-05: <https://workspaceupdates.googleblog.com/2024/05/tables-in-google-sheets.html>.
  - WU 2024-11: <https://workspaceupdates.googleblog.com/2024/11/table-reference-improvements-in-google-sheets.html>.
  - Release notes: <https://developers.google.com/workspace/sheets/release-notes> (2025-04-29: Tables in the Sheets API).
- **[measured]** is a live measurement the repo records, cited `file:line`. The paths are relative to `packages/framework/` unless they start with `byro-repo`.
- **[unverified]** means no primary source confirms it. It is an inference or a gap.

## The dry run is not the fake

This sits under every "fake models it?" answer below.

- **The chore dry run never replays requests onto the fake.** The Node host's dry run records each request, prints it through `UpdateRequestSummary`, and returns `{}`: `src/00_Source/GoogleSheets/GoogleSheetsAPI.ts:292-296` and `docs/how-it-runs/chore-dry-run.md:12-14`.
- **Reads in a dry run are live.** The same doc says so at line 12.
- **The fake is test-only.** It lives at `src/testSupport/fakeSheetsService.ts` with its replays in `src/testSupport/fakeSheetsService/`. With `isDryRun: true` the fake also "counts batch updates without replaying them" (`docs/testing.md:27`, `src/testSupport/fakeSheetsService.ts:397`).
- So "the fake models it" means "a test can assert the grid state it leaves". A dry-run preview gets nothing from the fake unless the design makes the fake replay against a snapshot of the live fetch.
- **The fake throws on what it doesn't model.** It throws on any request kind it doesn't replay (`src/testSupport/fakeSheetsService.ts:466-468`). It throws on any `updateTable` field outside `name`, `range` and `columnProperties` (`src/testSupport/fakeSheetsService/tableReplays.ts:47-57`, `fieldMasks.ts:7-14`). It throws on any cell field outside `userEnteredValue`, `userEnteredFormat(.backgroundColor|.numberFormat)` and `dataValidation` (`fakeCells.ts:10-16, 33-45`).

### What a preview needs before it can classify "touches typed data"

- **The read mask decides it.** The fetch that carries "programmatic facts" asks for `userEnteredValue`, so it can tell a formula cell (`userEnteredValue.formulaValue`) from a typed value (`stringValue`/`numberValue`/`boolValue`) and from a blank (`src/00_Source/GoogleSheets/GoogleSheetsAPI.ts:91-97`). [doc] `CellData.userEnteredValue` is "the value the user entered… e.g. 1234, 'Hello', or =NOW()" (Cells ref).
- **The standing fetch can't tell them apart.** It asks for `effectiveValue` only (`GoogleSheetsAPI.ts:98-104`), so a formula and a typed value of the same result look the same.
- A classifier needs the programmatic-facts mask over every cell a request can overwrite or delete.

## Summary table

| # | Property | Request that sets it | Framework emits it? | Fake models it? | Preview can flag typed-data loss? |
|---|---|---|---|---|---|
| 1 | Column type | `updateTable`, `fields: "columnProperties"`: a full-list replace | Yes: `setTableColumnProperties` (`GoogleSheetsAPI.ts:491-494, 513-525`), guarded (`TableRaw.ts:705-738`) | Yes: full-list replace, plus a refusal of a column with no name (`tableReplays.ts:62-118`) | Values are not deleted. The loss is the validation rules and dropdown styling, which a preview can flag from the fetched `columnProperties` |
| 2 | Column formula | No Table-level column formula exists. Write per cell: `repeatCell`/`updateCells` with `userEnteredValue.formulaValue`, or `pasteData` `PASTE_FORMULA` | Yes: `fillColumn` gives a formula as `pasteData` and a value or colour as `repeatCell` (`cellData.ts:40-115`, `TableRaw.ts:471`) | Yes: `pasteData` NORMAL/FORMULA and `repeatCell` (`cellReplays.ts:35-85`). No formula evaluation | Yes, with a `userEnteredValue` fetch of the target cells |
| 3 | Checkbox on an action-row cell | `setDataValidation` with `condition.type: BOOLEAN` | Yes: `addCheckboxValidation` (`GoogleSheetsAPI.ts:495-503`, `TableRaw.ts:517-531`, `CellRaw.ts:145`) | Yes (`cellReplays.ts:86-103`) | Not needed: the value is untouched. A preview can flag a typed non-boolean value it would mark invalid |
| 4 | Head-row background | `updateCells`/`repeatCell`, `fields: userEnteredFormat.backgroundColor` (or `...backgroundColorStyle`) | Yes: `fillCell` gives `updateCells` (`cellData.ts:52-65, 117-143`, `CellRaw.ts:135`) | Yes, `backgroundColor` only, not `backgroundColorStyle` (`fakeCells.ts:10-16`) | Not needed: format only |
| 5 | Header text | Writing the header cell (`updateCells`), or `columnName` in an `updateTable` `columnProperties` replace | Yes: header writes through `fillCell`, and `columnName` rides every column-type replace (`TableRaw.ts:740-758`) | Yes: a sent `columnName` writes the header cell (`tableReplays.ts:79-85`) | Not applicable to data. It may break table references in formulas [unverified] |
| 6 | Insert a column inside a Table | `insertRange` COLUMNS over the Table's rows (plus head rows), or `insertDimension` COLUMNS | Only at the Table's end: `insertRange` plus an `updateTable` range widen (`TableRaw.ts:570-600`). There is no mid-Table insert (`docs/architecture/table-columns.md:9`) | Yes: shift, growth, a refusal of a part-of-Table cut, and `Column <n>` headers (`fakeGrid.ts:104-176, 202-228, 423-438`) | Nothing is overwritten. The cells shift, so a preview can name what moves |
| 7 | Delete a Table column | `deleteRange` COLUMNS over the Table's rows (plus head rows), or `deleteDimension` COLUMNS | No column-delete path. `deleteRange` is only used for rows (`GoogleSheetsAPI.ts:382-387`) | Yes: `deleteRange`/`deleteDimension` shrink the Table (`rangeReplays.ts:33-43`, `fakeGrid.ts:142-156, 177-189`) | Yes, with a `userEnteredValue` fetch of the column |
| 8 | ID-prefix rename | `findReplace`, scoped to exactly one of `range`/`sheetId`/`allSheets` | Yes: column-scoped (`TableRaw.ts:548-558`, `ColumnRaw.ts:170-179`) and caller-scoped (`SpreadsheetRaw.ts:156-163`) | Partly: string cells only, regex via JS not Java, numbers never match (`cellReplays.ts:104-136, 207-234`) | Only from the read. The response counts (`valuesChanged`, `formulasChanged`) never reach a dry run (`UpdateRequestSummary.ts:183`) |
| 9 | Protected range | `addProtectedRange` (`range` or `tableId` or `namedRangeId`, `warningOnly`, `unprotectedRanges`) | Yes, by `range` only, never `tableId` (`TableRaw.ts:539-546`, `GoogleSheetsAPI/protectedRanges.ts:34-40`) | Yes: add, delete and shift (`ruleReplays.ts:31-65`, `fakeGrid.ts:451-457`) | Not needed: no data change |
| 10 | Conditional format on a column | `addConditionalFormatRule` with GridRange `ranges` only, no Table reference | Yes: index 0, ranges from the Table origin at gathering (`TableRaw.ts:532-538`) | Yes: add and delete by index, shifting on inserts and deletes, extending on a `PASTE_FORMAT` below (`ruleReplays.ts:6-30`, `rangeReplays.ts:129-171`) | Not needed: no data change |

## 1. Column type

**Request.** `updateTable { table: { tableId, columnProperties: [...] }, fields: "columnProperties" }`.
- [doc] `UpdateTableRequest` takes `table` and a required `fields` mask (Request ref).
- [doc] `TableColumnProperties` = `columnIndex` (relative to the Table), `columnName`, `columnType`, and `dataValidationRule`, which is "Only set for dropdown column type" (Sheets ref).
- [doc] The `ColumnType` enum: `COLUMN_TYPE_UNSPECIFIED, DOUBLE, CURRENCY, PERCENT, DATE, TIME, DATE_TIME, TEXT, BOOLEAN, DROPDOWN, FILES_CHIP, PEOPLE_CHIP, FINANCE_CHIP, PLACE_CHIP, RATINGS_CHIP` (Sheets ref).
- [doc] A dropdown column "must be set with a `ONE_OF_LIST` condition. Other column types shouldn't set the `dataValidationRule` field" (Tables guide).
- [doc] `TableColumnDataValidationRule` has one field, `condition`, "Valid only if the [BooleanCondition.type] is ONE_OF_LIST" (Sheets ref).

**What it does to neighbouring state.**
- **The list is replaced whole.** [measured] A `columnProperties` list replaces the Table's whole list. A column left out keeps its header cell but loses its type and validation. A sent `columnName` is written to the header cell (`docs/testing.md:33`, `src/testSupport/fakeSheetsService/tableReplays.ts:62`).
- [doc] This fits the field-mask rule: a field in the mask that is absent from the message is cleared (Field masks guide, "Update with a field mask"). The guide says nothing about addressing one element of a repeated field. [unverified] that a mask such as `columnProperties.columnType` can target one column. The repo never tried it, and the guide gives no syntax for list elements.
- **So every column-type change rewrites every column on the Table**: names, types and dropdown rules. The framework sends every column's `columnName` and `columnType` (`src/02_SpreadsheetRaw/TableRaw.ts:740-758`) but no `dataValidationRule`. That is one more reason it must refuse on a validated Table, since a resend without the rule would drop it ([measured] behaviour of an omitted rule, `docs/testing.md:33`).

**The dropdown style and colour trap, confirmed and generalised.**
- [doc] Neither validation schema has a field for a dropdown's display style (chip or arrow) or its per-option colours:
  - `TableColumnDataValidationRule` carries only `condition` (Sheets ref).
  - Cell-level `DataValidationRule` carries `condition`, `inputMessage`, `strict` and `showCustomUi` (Cells ref).
  - `ConditionValue` carries `userEnteredValue` / `relativeDate` (Other ref).
- **The API can't write those settings back.** Any `updateTable` that replaces `columnProperties` can only resend the rule's condition. Even a byte-identical resend of the fetched list can't carry the chip style or option colours.
- **Repo status.**
  - The claim that this reset *happens* is recorded as a design premise in byro-repo#80: "The API has no field for a dropdown's display style (arrow or chip) or its option colours, so any such update resets them on every dropdown in that Table".
  - The guard is `TableRaw.ts:722-731`. The docs are `docs/architecture/table-columns.md:13` and `docs/generated-data/config-sheet-floor.md` ("What every config sync repairs").
  - **[unverified]: no live measurement of the reset itself is recorded.** The issue states it as an inference from the schema. A one-off scratch-tab probe would settle it. Two cases need checking: whether colours survive a resend of an identical `columnProperties` list, and whether they survive on a dropdown column whose entry is resent unchanged while only another column's type changes.
- **Generalised:** any write whose mask includes `columnProperties` puts every dropdown column on that Table at risk. That covers a column type, a header rename through `columnName`, and `addTable` with column properties. Writes masked to `name`, `range` or `rowsProperties` don't touch `columnProperties`, so the dropdowns aren't at risk from them ([doc] the mask "ignores any fields that aren't specified", Field masks guide). Whether a `range`-only widen leaves dropdown colours alone is [unverified] live. The repo widens on every growth (`GoogleSheetsAPI.ts:343-351`) with no reported colour loss.
- **Related quirk.** [measured] Column names sent inside `addTable` "land three columns right", so the adapter sends `addTable` without `columnProperties` and follows it with an `updateTable` (`GoogleSheetsAPI.ts:453-466`, `tableReplays.ts:16-26`).

**Refusals.**
- [measured] A sent column with no `columnName` is rejected (`docs/testing.md:33`, mirrored at `tableReplays.ts:100-105`).
- [doc] `DROPDOWN` without a `ONE_OF_LIST` rule, or a rule on a non-dropdown type, violates the guide's "must" / "shouldn't" (Tables guide). The exact error text is [unverified].
- [unverified] Whether the API accepts a `ONE_OF_RANGE` rule on a `DROPDOWN` column. The schema says only `ONE_OF_LIST` is valid, yet the app's dropdowns use `=<Table>[<Header>]` range sources (`docs/generated-data/column-configs.md:29`), which may arrive with `columnType` unset rather than `DROPDOWN`.

**Framework guards** (`TableRaw.ts:694-738`):
- one update per fetch;
- a refusal when the same flush inserts a column on that sheet;
- a refusal when any fetched column has a validation condition or values;
- a refusal when any column has no fetched name.

**Fake.** It models the full-list replace and the missing-name refusal. It throws on any rule field other than `condition.type` / `values[].userEnteredValue` (`tableReplays.ts:100-118`). It has no notion of dropdown style or colours, so a test can't catch the reset.

**Preview.** No cell value changes. What is lost is the rule and styling on other columns, and a preview can detect that before sending: the fetched `columnProperties` (`GoogleSheetsAPI.ts:94`) show which columns carry a `dataValidationRule`. [doc] Typed values that don't fit the new type are kept and shown with a warning: "Data entered that does not align with a set column type will result in a warning" (WU 2024-05; Help: tables, "Set column types"). [unverified] whether `columnType` also reformats existing cells. [doc] The rating and checkbox types "populate with default values of 0 and FALSE" (Tables guide). Whether that fills blank existing cells on a type change is [unverified]. If it does, a type change *writes* values.

## 2. A formula across a Table column

**No Table-level column formula exists in the API.**
- [doc] `Table` has `tableId`, `name`, `range`, `rowsProperties` and `columnProperties`. `TableColumnProperties` has no formula field (Sheets ref).
- [doc] The Tables guide and the help pages describe no calculated-column or auto-fill behaviour for Tables.
- [doc] Table references exist (`Table1[Column 1]`, `#ALL`, `#HEADERS`, `#TOTALS`, `#DATA`), and "#This Row currently is not supported" (Help: table refs; WU 2024-11). So a per-row formula can't use Excel's `[@Col]` idiom.
- [doc] "use ARRAYFORMULA" to get values over a whole Table column (Help: table refs). A single `ARRAYFORMULA` in the first data row is the only way to get one formula covering the column. It then fails with `#REF!` when any typed value sits below it ([unverified]: general Sheets spill behaviour, not checked against a Google page here).
- [unverified] Whether a new row inserted inside a Table copies the formula from its neighbours. [measured] for rows, "an insert strictly inside a Table gives its new cells the format and validation of the cells below" (`docs/testing.md:40`), but no value or formula copy is recorded. The framework's growth copies format and validation only (`TableRaw.ts:626-648`).

**Which request writes it.**
- `repeatCell` with `cell.userEnteredValue.formulaValue` and `fields: "userEnteredValue"`. [doc] "If writing a cell with a formula, the formula's ranges will automatically increment for each field in the range… To keep the formula's ranges static, use the `$` indicator" (Request ref, `RepeatCellRequest`).
- `updateCells`, one `CellData` per row. [doc] With `range` set, "If the data in rows does not cover the entire requested range, the fields matching those set in `fields` will be cleared" (Request ref).
- `pasteData` with `type: PASTE_FORMULA` (Request ref). The framework uses this: each row gets the same formula text (`src/00_Source/GoogleSheets/GoogleSheetsAPI/cellData.ts:98-115`). The comment at `TableRaw.ts:471` gives the reason: "a formula fill is pasteData so Sheets parses it". [unverified] whether `pasteData` shifts relative references between pasted rows the way `repeatCell` does. The framework pastes identical text, so a relative A1 reference would point at the same cell in every row unless pasteData adjusts it.

**Neighbouring state.**
- Every cell in the run has its `userEnteredValue` replaced, typed values included.
- Format is untouched unless it is in the mask.

**Refusals.**
- [measured] A range starting past the grid edge is refused, and one running past it is clipped (`docs/testing.md:43`; measured for `insertRange`/`deleteRange`/`copyPaste`, [unverified] for `repeatCell`/`pasteData`).
- A malformed formula is stored as a formula that evaluates to an error, not refused ([unverified]).

**Framework.** `fillColumn` gives `pasteData` for a formula, then `repeatCell` for a value or colour (`cellData.ts:40-51, 68-86`). It also has `updateWorkingFormulas` (`ColumnRaw.ts:161-168`).

**Fake.** It replays `pasteData` NORMAL/FORMULA and stores the formula text without evaluating it (`cellReplays.ts:49-85`). It replays `repeatCell` and `updateCells` with field masks (`cellReplays.ts:14-48`). It does not model the reference incrementing.

**Preview.** Yes, it can tell. Every target cell is known from the request's range or coordinate and row count, and a programmatic-facts fetch says whether each holds a formula, a typed value or nothing.

## 3. A checkbox on an action-row cell above the header

**Request.** `setDataValidation { range, rule: { condition: { type: "BOOLEAN" } } }`.
- [doc] It "Sets a data validation rule to every cell in the range. To clear validation in a range, call this with no rule specified" (Request ref).
- [doc] `BOOLEAN`: "The cell's value must be TRUE/FALSE or in the list of condition values… Renders as a cell checkbox. Supports zero, one or two ConditionValues" (Other ref).
- An alternative is `updateCells`/`repeatCell` with `fields: "dataValidation"` (Cells ref, `CellData.dataValidation`).

**Neighbouring state.**
- It replaces the cell's existing validation, whatever kind it was, and leaves the value and format alone.
- A cell outside the Table range is ordinary grid: no Table column type applies.
- [measured] A column insert that inherits from before copies its left neighbour's cell validation into every row, "above the Table too", and the copied checkbox arrives unticked (`docs/testing.md:32`). The framework's own end insert does not inherit, and its new cells are blank (`docs/testing.md:41`).
- [measured] The framework's growth copies cell validation only on untyped columns, with `PASTE_DATA_VALIDATION` (`TableRaw.ts:640-647`). This is irrelevant above the header.

**Refusals.**
- None documented for a valid single-cell range, beyond the grid-limit rule (`docs/testing.md:43`).
- [unverified] What a non-boolean typed value already in the cell shows: it is probably flagged invalid and kept, since `strict` defaults to false (Cells ref, `strict`: "True if invalid data should be rejected").

**Framework.** `CellRaw.addCheckboxValidation` (`CellRaw.ts:145`) and `TableRaw.addCheckboxValidationAt` (`TableRaw.ts:491-531`) give `addCheckboxValidation`, then `setDataValidation` (`GoogleSheetsAPI.ts:495-503`). Also `SpreadsheetRaw.ts:150`.

**Fake.** Yes (`cellReplays.ts:86-103`). It stores the rule; it does not seed `FALSE`.

**Preview.** The value is untouched. A preview could warn when a typed value other than TRUE/FALSE is present.

## 4. Background colour of head rows above the Table header

**Request.** `updateCells` (one cell) or `repeatCell` (a run) with `cell.userEnteredFormat.backgroundColorStyle` (or the deprecated `backgroundColor`) and a matching mask.
- [doc] `CellFormat.backgroundColor` is "(deprecated)… Use backgroundColorStyle". `backgroundColorStyle`: "If backgroundColor is also set, this field takes precedence" (Cells ref).
- The Table's *own* header row colour is a different property: `Table.rowsProperties.headerColorStyle`, "If this field is set, the header row is filled with the specified color. Otherwise… a default color" (Sheets ref). That is set by `updateTable` with `fields: "rowsProperties.headerColorStyle"` ([unverified] that this nested mask path is accepted; the Field masks guide gives only the generic syntax). The head rows above the header are outside the Table range, so `rowsProperties` does not reach them.

**Neighbouring state.**
- Only the masked format field changes.
- [unverified] **Trap: the framework writes the deprecated `userEnteredFormat.backgroundColor` only** (`cellData.ts:117-143`). If a cell already carries a `backgroundColorStyle` (for example a theme colour set in the UI), the documented precedence implies the style wins and the write has no visible effect. Masking `backgroundColorStyle` instead avoids that.
- Inside the Table range, banding colours (`firstBandColorStyle`/`secondBandColorStyle`) also colour cells. Which wins over a cell background there is [unverified].

**Refusals.** None documented.

**Framework.** `CellRaw.updateBackgroundColor` (`CellRaw.ts:135-143`) gives `fillCell`, then `updateCells` (`cellData.ts:52-65`). A column fill gives `repeatCell` (`cellData.ts:40-51`).

**Fake.** It models `userEnteredFormat.backgroundColor` and throws on `backgroundColorStyle` (`fakeCells.ts:10-16, 33-45`). It does not model `rowsProperties` (`tableReplays.ts:47-57` throws on that field).

**Preview.** Format only; no data is touched.

## 5. Header text of a Table column

**Requests.** There are two paths, and the repo has measured both:
- **Writing the header cell's value** (`updateCells` on the Table's first row) renames the column. [measured] Header writes in the same flush as a `columnProperties` update are sent *after* it "so a header write in the same flush renames the column rather than being reverted" (byro-repo#80 acceptance criteria). The sync also overwrites floor headers back to the seed (`docs/generated-data/config-sheet-floor.md`, "What every config sync repairs").
- **`updateTable` `columnProperties[].columnName`**. [measured] "a sent `columnName` is written to the header cell" (`docs/testing.md:33`). This is a full-list replace, so section 1's dropdown trap applies.

**Neighbouring state.**
- [doc] Table references name columns by header (`Table1[Column 1]`) and "update whenever you add or remove data" (Help: table refs). Whether existing formulas that use a reference are rewritten when the header is renamed is [unverified]. The help pages state only that references track data changes, not renames.
- The framework's dropdown sources are `=<Table>[<Header>]` (`docs/generated-data/column-configs.md:29`), so a rename may break or silently retarget them ([unverified]).

**Refusals.**
- [measured] An empty or missing `columnName` in `columnProperties` is rejected (`docs/testing.md:33`).
- [unverified] Duplicate header names within a Table, and blank header cells written through `updateCells`.
- [doc] Table *names* (not headers) can't be TRUE/FALSE, look like A1/R1C1, start with a number, exceed 255 characters, or hold special characters other than underscores (Help: tables). No equivalent rules for headers are documented.

**Framework.**
- Header cells are written by `fillCell` (for example the end-column insert writes "column ID, header and optional group heading": `docs/architecture/table-columns.md:9`).
- `columnName` rides every column-type replace (`TableRaw.ts:740-758`).
- Table renames are `updateTable` `fields: "name"` (`GoogleSheetsAPI.ts:479-490`).

**Fake.** Yes, both paths (`tableReplays.ts:79-85`, `cellReplays.ts:14-34`).

**Preview.** No data rows are touched.

## 6. Inserting a column at an index inside a Table

**Requests.**
- [doc] "If you need to add a new row or column within the table, use the InsertRangeRequest or the InsertDimensionRequest… If using InsertRangeRequest/DeleteRangeRequest, the range in the request must cover the entire row(s) if deleting a row or entire column(s) if deleting a column" (Tables guide).
- [doc] `insertDimension` inserts whole sheet columns. `inheritFromBefore` decides whose "dimension properties" (for example background) the new columns take, and must be false at index 0 (Request ref).
- [doc] `insertRange` "Inserts cells into a range, shifting the existing cells over or down", constrained to the sheet's bounds (Request ref).

**Neighbouring state** (all [measured], `docs/testing.md:32, 40-42`):
- **Growth.** An insert strictly inside a Table grows it. One just past its last column does not grow it unless it is `insertDimension` with `inheritFromBefore: true`. Hence ADR 0013's insert-plus-widen for the end case (`docs/adr/0013-grow-a-table-by-insert-plus-widen.md`, `TableRaw.ts:570-600`).
- **Which cells move.** `insertRange` shifts only the cells inside the band, so a Table beside it is spared. `insertDimension` shifts the whole sheet column, every other Table and cell in it included.
- **New headers.** A column a Table grows by is headed `Column <n>` (`fakeGrid.ts:423-438`), and its Table column type starts Automatic (`docs/architecture/table-columns.md:9`, for the end case).
- **Format of the new cells.** The inheriting `insertDimension` copies the left neighbour's cell format and validation into every row (including above the Table), but no Table column type or column rule. An end-column `insertRange` leaves the new cells blank.
- **What shifts with it.** Columns to the right shift, along with their column-keyed facts (types, rules) (`fakeGrid.ts:407-421`), protected ranges, conditional-format ranges and hidden indexes.
- [unverified] The format and validation of a column inserted *strictly inside* a Table. Only rows were measured there ("the format and validation of the cells below"), and the fake leaves them blank (`docs/testing.md:40`).
- [unverified] Whether table references and formulas pointing at columns to the right are rewritten.
- The head rows (the column-ID row at header −3, and so on) sit above the Table. An `insertRange` must span them too, or they stop lining up with their columns (`docs/architecture/table-columns.md:9`).

**Refusals** ([measured]):
- `You cannot insert or delete cells over part of a table.` This happens when a Table reaches into the band's shifted span but only part of it lies across the band (`docs/testing.md:42`; adapter detection at `GoogleSheetsAPI.ts:253-266`).
- `Range (...) exceeds grid limits` when the range starts past the edge; it is clipped when it runs past (`docs/testing.md:43`).
- The framework refuses a column type change in the same flush as a column insert on that sheet (`TableRaw.ts:713-720`). It also refuses a neighbouring Table split by the band (`TableRaw.ts:601-609`).

**Framework.** Only an end-of-Table insert exists (`TableRaw.appendColumn`, which gathers into `insertRange` COLUMNS plus `updateTableRange`, with `appendDimension` first when needed: `docs/architecture/table-columns.md:9`). There is no mid-Table insert (same doc). `insertDimension` is in the adapter grammar but has no Raw caller.

**Fake.** Yes: growth rules, the part-of-Table refusal, `Column <n>` headers, shifts of rules and protections (`fakeGrid.ts:104-176, 202-228`). It leaves the new mid-Table cells blank, which is unmeasured live. A rule or protection cut across throws as unmodelled (`docs/testing.md:42`).

**Preview.** Nothing is overwritten. A preview can list every cell, rule and protection that shifts, plus every Table on the sheet a `insertDimension` would also cut through.

## 7. Deleting a Table column

**Requests.**
- [doc] `deleteDimension` (whole sheet columns) or `deleteRange` with `shiftDimension: COLUMNS` "shifting other cells into the deleted area" (Request ref).
- [doc] The Tables guide's note that the range must "cover… entire column(s)" applies.
- [doc] `deleteTable` "Removes the table with the given ID… and the contents of the table". Deleting banding instead removes the Table formatting but keeps the data (Tables guide).

**Neighbouring state.**
- [measured] "A delete inside a Table shrinks it" (`docs/testing.md:32`).
- `deleteDimension` removes the column for the whole sheet, so every other Table or cell in it goes too.
- [measured] `deleteRange` shifts only the band's cells and leaves the grid size alone (`docs/testing.md:40`).
- The column's values, type, rule and any conditional-format or protected range sitting wholly inside it are removed (`fakeGrid.ts:407-421, 451-465`; [measured] for shifts, `docs/testing.md:32`).
- [unverified] Formulas elsewhere that reference the column become `#REF!`, and table references to a deleted header break. This is general Sheets behaviour with no Google page checked here.

**Refusals.**
- [measured] Part-of-Table cuts (as in section 6).
- [measured] "Cannot delete a table header row" applies to row deletes (`fakeGrid.ts:213-228`, `docs/testing.md:40`); it is not relevant to a column.
- [unverified] What happens on deleting a Table's last remaining column.

**Framework.** There is no column-delete path. `deleteRange` is emitted only for row deletes (`deleteTableRows`, `GoogleSheetsAPI.ts:382-387`), and the blank-row deletion paths are covered by `docs/architecture/blank-row.md`. The `deleteDimension` replay exists for raw requests only (`dimensionReplays.ts:74-83`).

**Fake.** Yes for both kinds (`rangeReplays.ts:33-43`, `fakeGrid.ts:142-156, 177-189`).

**Preview.** Yes, and this is the most important case. Every cell in the deleted band is known from the request, and a programmatic-facts fetch shows how many hold typed values, and which, as opposed to formulas or blanks.

## 8. Renaming an ID prefix with `findReplace`

**Request.** `findReplace { find, replacement, matchCase, matchEntireCell, searchByRegex, includeFormulas, range | sheetId | allSheets }`. [doc] (Request ref):
- **Scope:** exactly one of `range`, `sheetId` or `allSheets`.
- **Regex:** `searchByRegex` follows Java regex rules, and "The replacement string is allowed to refer to capturing groups" (`$1`).
- **Formulas:** `includeFormulas`: "True if the search should include cells with formulas. False to skip cells with formulas".
- **Response:** `valuesChanged` ("non-formula cells changed"), `formulasChanged`, `rowsChanged`, `sheetsChanged` and `occurrencesChanged` (Response ref).

**What it touches.**
- Cell contents only, over the scope.
- With `includeFormulas: true` it rewrites the formula *text*, so a prefix inside a string literal or a reference in a formula is changed too.
- A scope that includes the header row renames columns (section 5).
- [unverified] Whether it matches the stored value or the displayed text, and whether number or date cells can match. The fake treats only strings as matchable (`cellReplays.ts:104-121`), which is unmeasured.
- [unverified] It does not touch `ONE_OF_LIST` option values in validation rules, conditional-format formulas, named ranges or Table names, since those are not cell contents. So an ID used as a dropdown option would leave cells invalid until the rule is rewritten.
- [unverified] Whether `^` anchors at the start of each cell under `searchByRegex`. `matchEntireCell` is the documented whole-cell anchor.
- **Trap:** the ID prefix also appears in other Tables' reference columns (IDs as foreign keys). A column-scoped rename leaves those stale, and a sheet-wide or `allSheets` rename reaches typed values in every Table.

**Refusals.** Setting more than one scope ([doc] "one and only one must be set"). Error text [unverified].

**Framework.**
- `ColumnRaw.findReplace` (`ColumnRaw.ts:170-179`) and the Table-level one are scoped to the body rows of one column run (`TableRaw.ts:547-558`).
- `SpreadsheetRaw.findReplace` takes a caller scope (`SpreadsheetRaw.ts:156-163`).
- The preview prints "matching cells" with no count, because "The count lives in a response a dry run never gets" (`src/nodeHost/UpdateRequestSummary.ts:183-192`).

**Fake.** It replays `allSheets`, `sheetId` and `range` scopes, `matchCase`, `matchEntireCell`, `includeFormulas`, and regex through JavaScript `RegExp` rather than Java (`cellReplays.ts:104-136, 207-234`). The differences between JS and Java regex (for example the `$` group syntax, possessive quantifiers) are not modelled.

**Preview.** Only by matching the pattern itself against a programmatic-facts fetch of the scope, which splits matches into formula cells and typed cells. The API's own `valuesChanged` / `formulasChanged` split arrives only after a real send.

## 9. Protected ranges

**Request.** `addProtectedRange { protectedRange: { range | namedRangeId | tableId, description, warningOnly, unprotectedRanges, editors } }`. [doc] (Sheets ref `ProtectedRange`, Request ref):
- **Target:** "When writing, only one of range or namedRangeId or tableId may be set". `tableId`: "The table this protected range is backed by". The Tables guide lists protected ranges among the features that "support tables being their backing data".
- **`warningOnly`:** "every user can edit… except editing will prompt a warning". When it is true, `editors` are ignored. Changing it from true to false without `editors` makes every document editor an editor.
- **`unprotectedRanges`:** "only supported on protected sheets", meaning a fully unbounded `range`.
- **IDs:** `protectedRangeId` is optional on add, and "It is an error to specify the ID of a range that already exists".
- **`requestingUserCanEdit`:** read-only.

**Neighbouring state.**
- None on cells.
- [measured] A range-backed protection shifts with inserts and deletes. One the band cuts across is unmodelled in the fake, and a whole-sheet protection stays as it is (`docs/testing.md:32, 42`).
- [unverified] Whether a `tableId`-backed protection follows the Table as it grows (likely, given "backed by", but no measurement or doc statement).

**Refusals.**
- A duplicate ID ([doc]).
- `unprotectedRanges` on a range that is not a whole sheet ([doc] "only supported on protected sheets"; the error text is [unverified]).
- [measured] `deleteProtectedRange` with an unknown id: `No protected range with id` (`docs/testing.md:35`).

**Framework.** Range-backed only, never `tableId`. `warningOnly` comes from `protection.kind === "warning"`, and `unprotectedRanges` passes through (`GoogleSheetsAPI/protectedRanges.ts:34-40`). The protection is queued from `TableRaw.gatherEditProtectionOperation` (`TableRaw.ts:539-546`), from sheet-level edit protections (`SheetRaw/SheetEditProtectionsRaw.ts:71-88, 139-140`), and from the floor's "one whole-sheet edit warning with editable ranges" (`docs/generated-data/config-sheet-floor.md`).

**Fake.** Add, delete, shift and a sequential ID (`ruleReplays.ts:31-65`, `fakeGrid.ts:451-457`).

**Preview.** No data change.

## 10. Conditional formatting scoped to a Table column

**Request.** `addConditionalFormatRule { rule: { ranges: [GridRange], booleanRule | gradientRule }, index }`.
- [doc] "Adds a new conditional format rule at the given index. All subsequent rules' indexes are incremented" (Request ref).
- [doc] `ConditionalFormatRule.ranges[]` are GridRanges, "All the ranges must be on the same grid". There is no `tableId` (Sheets ref).
- [doc] **A rule can't be scoped to a Table by reference**: "These features don't currently support the use of table references when you select a range: Conditional formatting, Charts, Pivot tables" (Help: table refs). Whether a `CUSTOM_FORMULA` condition may *contain* a table reference is [unverified].
- `updateConditionalFormatRule` replaces or moves the rule at an index. `deleteConditionalFormatRule` takes `sheetId` + `index` (Request ref).

**Does it track the Table?** Only through ordinary grid-range adjustment, never through the Table's identity:
- [measured] Rule ranges shift with inserts and deletes (`docs/testing.md:32`).
- [measured] A `PASTE_FORMAT` onto the rows just below a rule's range extends the rule (`docs/testing.md:44`, `rangeReplays.ts:129-171`, sheets-framework#61). That is how the framework's growth keeps rules covering new rows.
- [unverified] Whether `updateTable` `range` widening or an `appendCells` with `tableId` extends a rule.
- **Trap:** rules are addressed by index per sheet. Adding at index 0 shifts every other rule's index, so a delete-by-index must use indexes from a fresh read taken after any adds in the same batch.

**Refusals.**
- [unverified] Exact errors. Ranges on different sheets are invalid by [doc] ("must be on the same grid").
- [measured] `deleteConditionalFormatRule` at a missing index is refused (mirrored at `ruleReplays.ts:16-30`).

**Framework.** `TableRaw.gatherConditionalFormatRuleOperation` adds at index 0, with ranges computed from the Table origin at gathering (`TableRaw.ts:532-538`). Also `ColumnRaw.ts:204` and `CellNamed`/`ColumnNamed`/`TableNamed` (`04_SpreadsheetNamed/*`). Sheet-level delete and re-add lives in `SheetRaw/SheetConditionalFormatsRaw.ts:41-67` and is flushed in `SpreadsheetFlusherRaw.ts:147-154, 193, 293`.

**Fake.** Add and delete by index, shift on insert and delete, extend on a format paste below (`ruleReplays.ts:6-30`, `fakeGrid.ts:458-465`, `rangeReplays.ts:129-171`).

**Preview.** No data change.

## Traps for the declarations design

1. **A `columnProperties` write is all-or-nothing per Table.** A declared type, header or dropdown change on one column rewrites every column's name, type and rule, and can't restore dropdown chip style or option colours ([doc]: no such fields). Treat any declaration that needs `columnProperties` as a whole-Table change. Refuse it, or require an explicit acknowledgement, on any Table with a dropdown. The live reset of an *identical* resend is still [unverified], so a scratch-tab probe is worth running before relying on either answer.
2. **No column formula exists.** A declared column formula has to be written per cell (`pasteData`/`repeatCell`). It does not auto-fill new rows (ADR 0013 growth copies only format and validation), and `#This Row` is unsupported. So the declaration needs a regen-time re-fill, or an `ARRAYFORMULA` in the first row.
3. **`insertDimension`/`deleteDimension` cut the whole sheet.** Every other Table in those columns is hit. Prefer `insertRange`/`deleteRange` bounded to the Table plus its head rows. Expect the `part of a table` refusal when a neighbouring Table overlaps the band partly.
4. **Conditional formats and protections are GridRange-scoped.** They follow inserts, deletes and the growth `PASTE_FORMAT`, never the Table itself. Protections *can* be `tableId`-backed ([doc]), which the framework doesn't use yet. Conditional formats can't be ([doc]).
5. **`findReplace`'s data-loss signal arrives too late.** `valuesChanged` versus `formulasChanged` arrives only in the response, so a dry run can't report it. The scope is content-matched, so a preview must run the pattern itself over a `userEnteredValue` fetch. Java regex differs from the fake's JS regex.
6. **The dry run is a recorder over live reads, not a fake replay.** Classifying "deletes or overwrites a typed value" needs the programmatic-facts fetch (`userEnteredValue`) over the target cells. The standing fetch's `effectiveValue` can't tell a formula from a typed value.
7. **`backgroundColor` versus `backgroundColorStyle`.** The framework writes the deprecated field, and the documented precedence means an existing `backgroundColorStyle` could mask it ([unverified] live).

## Sources

- Google: the Tables guide, the Sheets/Request/Cells/Other/Response REST references, the field-masks guide and the release notes under developers.google.com/workspace/sheets. Help Center 14239833 and 15637642. Workspace Updates 2024-05 (tables launch) and 2024-11 (table reference improvements). All were fetched 2026-10-31; URLs are listed at the top.
- Repo: the `file:line` citations above. Also byronbroughten/byro-repo#80 (the issue that introduced the column-type guard; it states the dropdown reset as reasoning from the schema).
