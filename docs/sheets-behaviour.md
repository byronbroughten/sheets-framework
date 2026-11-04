# Sheets behaviour

An [observation index](https://github.com/byronbroughten/byro-repo/blob/master/docs/agents/prose-files.md#terms) of how Google's platforms behave (the Sheets UI, the Sheets REST API, Apps Script with `SpreadsheetApp` and triggers, and the Google auth they share), measured costs included, never the framework's own behaviour. It is grouped by subject and follows the [observation-index rules](https://github.com/byronbroughten/byro-repo/blob/master/docs/prose-style.md#observation-indexes).

Whoever settles or overturns an observation adds or updates its line in the same change as the finding.

Write-ups sourced from Google's docs live on `research/*` branches, listed by `git branch -r --list 'origin/research/*'`.

## Tables and growth

### Appending rows to a Table

- Observed: REST API v4 `appendCells` with a `tableId` grew the grid by the rows it inserted, and `appendCells` with no `tableId` wrote after the last row holding a value; scratch tab on the dev spreadsheet; 2026-09-26; [recorded in db5ba74](https://github.com/byronbroughten/sheets-framework/commit/db5ba74d35cb24b9cf97975a1d86a589b88c7a76).
- Observed: REST API v4 `appendCells` keyed on a `tableId` inserted whole sheet rows below the Table's last row: a stray value below, a lower Table with its head rows, and a value in a column outside the Table all moved down by the appended count, and an empty row between the Table and the next occupied row stayed a gap; side by side with headers on one row, a shorter neighbour Table was untouched, a same-height one grew by a blank row, and a taller one got a blank row inside its body; scratch tabs on the dev spreadsheet; 2026-10-02; [resolution on #56](https://github.com/byronbroughten/sheets-framework/issues/56#issuecomment-5963055620).
- Observed: REST API v4 `appendCells` onto a Table whose only body row was blank wrote into that row, and nothing below moved; scratch tab on the dev spreadsheet; 2026-10-02; [resolution on #56](https://github.com/byronbroughten/sheets-framework/issues/56#issuecomment-5963055620).
- Observed: REST API v4 `appendCells` onto a header-only Table returned `Internal error encountered.`, for a Table shrunk to its header and for one created header only; scratch tabs on the dev spreadsheet; 2026-10-02; [resolution on #56](https://github.com/byronbroughten/sheets-framework/issues/56#issuecomment-5963055620).
- Observed: On the row REST API v4 `appendCells` added to a Table, Google filled the formula columns (an A1 formula and a `SINGLE` Table reference), and the row kept the column types, the alternating colours, the body cells' text formatting and background and an untyped column's cell-level checkbox validation; a conditional-format rule ending at the Table's last row grew from `A4:A6` to `A4:A8`; scratch tabs on the dev spreadsheet; 2026-10-02; [resolution on #59](https://github.com/byronbroughten/sheets-framework/issues/59#issuecomment-5963486654), [resolution on #61](https://github.com/byronbroughten/sheets-framework/issues/61#issuecomment-5963889082).
- Observed: In the Sheets UI, a row added to a Table by hand came with the Table's formula columns filled in; sheet not recorded; reported 2026-11-02; [developer, quoted in the resolution on #183](https://github.com/byronbroughten/sheets-framework/issues/183#issuecomment-6595101709).

### Growing a Table with an insert and a widen

- Observed: A REST API v4 `insertRange` (`ROWS`) directly below a Table's last row, bounded to its columns, left the Table's range alone, and an `updateTable` widening the range by the new rows, in the same batch or a later one, took them in; on a header-only Table and on Tables with body rows; scratch tabs on the dev spreadsheet; 2026-10-02; [resolution on #56](https://github.com/byronbroughten/sheets-framework/issues/56#issuecomment-5963055620), [resolution on #59](https://github.com/byronbroughten/sheets-framework/issues/59#issuecomment-5963486654).
- Observed: In REST API v4 growth by `insertRange` below the Table plus an `updateTable` widen, a stray value directly below in the Table's columns and a lower Table within its columns (with head rows, touching it, or narrower) moved down by the inserted count intact, a value in a column outside the Table stayed put, and a side-by-side neighbour of the same height or taller was untouched; two growths of one Table in one batch each saw the earlier one; scratch tabs on the dev spreadsheet; 2026-10-02; [resolution on #59](https://github.com/byronbroughten/sheets-framework/issues/59#issuecomment-5963486654).
- Observed: The same REST API v4 growth with a lower Table wider than the growing one was refused with `Invalid requests[0].insertRange: You cannot insert or delete cells over part of a table.`, and the batch changed nothing; scratch tab on the dev spreadsheet; 2026-10-02; [resolution on #59](https://github.com/byronbroughten/sheets-framework/issues/59#issuecomment-5963486654).
- Observed: The rows REST API v4 growth by `insertRange` plus an `updateTable` widen added kept the column types (`CURRENCY` and `NUMBER` formats, both `DROPDOWN`s, `BOOLEAN`), the alternating colours and a whole-column conditional-format rule, and had empty formula cells, none of the body cells' text formatting or background, and no cell-level checkbox validation; a conditional-format rule ending at the Table's last row stayed `A4:A6`; scratch tabs on the dev spreadsheet; 2026-10-02; [resolution on #59](https://github.com/byronbroughten/sheets-framework/issues/59#issuecomment-5963486654), [resolution on #61](https://github.com/byronbroughten/sheets-framework/issues/61#issuecomment-5963889082).

### Inserting and deleting cells beside and inside a Table

- Observed: A REST API v4 `insertRange` or `deleteRange` (`ROWS`) bounded to one Table's columns grew or shrank only that Table, leaving a side-by-side Table's range and values untouched; a `deleteDimension` of a sheet row, and a `deleteRange` across both Tables' columns, shrank both Tables; probe sheet on the dev spreadsheet; 2026-10-01; [resolution on #53](https://github.com/byronbroughten/sheets-framework/issues/53#issuecomment-5933531939).
- Observed: A REST API v4 `insertRange` or `deleteRange` shifted only the cells inside the range's band, a Table the band held whole moved with it, a Table grew by an insert strictly inside it and not by one just below its last row or just past its last column, and an insert strictly inside a Table gave its new cells the format and validation of the cells below; scratch tabs on the dev spreadsheet; 2026-10-02, rerun 2026-10-03; [recorded in 7ff741d](https://github.com/byronbroughten/sheets-framework/commit/7ff741d434fda5aa91d3763c5e5da711c75ddf27).
- Observed: A REST API v4 `deleteRange` (`ROWS`) bounded to a Table's columns removed its only body row, and both of two body rows, leaving the Table header only; scratch tabs on the dev spreadsheet; 2026-10-02; [resolution on #56](https://github.com/byronbroughten/sheets-framework/issues/56#issuecomment-5963055620).
- Observed: A REST API v4 `deleteRange` of a Table's header row was refused with `Cannot delete a table header row. Consider hiding the row instead.`; dev spreadsheet; 2026-10-01, rerun 2026-10-03; [resolution on #53](https://github.com/byronbroughten/sheets-framework/issues/53#issuecomment-5933531939), [recorded in 7ff741d](https://github.com/byronbroughten/sheets-framework/commit/7ff741d434fda5aa91d3763c5e5da711c75ddf27).
- Observed: A REST API v4 `insertRange` or `deleteRange` whose band cut part of a Table was refused with `Invalid requests[<i>].<kind>: You cannot insert or delete cells over part of a table.`, naming only the request; scratch tabs on the dev spreadsheet; 2026-10-03; [recorded in 7ff741d](https://github.com/byronbroughten/sheets-framework/commit/7ff741d434fda5aa91d3763c5e5da711c75ddf27).
- Observed: REST API v4 sheet row and column inserts and deletes shifted the Tables, their column types, conditional-format ranges and hidden indexes with the rows and columns; a column insert at a Table's end grew it only with `inheritFromBefore: true`, which headed a new column `Column <n>`, and a delete inside a Table shrank it; scratch tab on the dev spreadsheet; 2026-09-26; [recorded in db5ba74](https://github.com/byronbroughten/sheets-framework/commit/db5ba74d35cb24b9cf97975a1d86a589b88c7a76).
- Observed: A REST API v4 column insert with `inheritFromBefore: true` at a Table's end copied its left neighbour's cell format and cell validation into every row, above the Table too, a copied checkbox arriving unticked (`false`), and copied no Table column type or column validation rule, beside a checkbox, dropdown and currency column; scratch tab on the dev spreadsheet; 2026-09-26; [recorded in 98a56b4](https://github.com/byronbroughten/sheets-framework/commit/98a56b489b733aa931c1de779e497c1bfc972266).

### Column inserts bounded to a Table

- Observed: A REST API v4 `insertRange` (`COLUMNS`, rows from three above the header to the Table's last row) at a Table's last column widened the Table with the new column second to last, headed `Column 1`; just past the last column it left the Table's range alone and shifted the band's cells to the right; just past the last column with an `updateTable` widening the range in the same batch, the Table widened with the new column last, headed `Column 1`; cells above and below the band stayed put; scratch tabs on the dev spreadsheet; 2026-10-02; [resolution on #56](https://github.com/byronbroughten/sheets-framework/issues/56#issuecomment-5963055620).
- Observed: With a side-by-side neighbour Table on the same header row, that REST API v4 column insert moved a neighbour of the same height or shorter right whole, head rows included, and was refused with `Invalid requests[0].insertRange: You cannot insert or delete cells over part of a table.` for a taller neighbour, with or without its head rows, and for a neighbour whose Table range started above the band; scratch tabs on the dev spreadsheet; 2026-10-02; [resolution on #56](https://github.com/byronbroughten/sheets-framework/issues/56#issuecomment-5963055620).
- Observed: In the Sheets UI, a column insert inside the left of two side-by-side Tables with headers on one row went through when the right Table was the same height or shorter, and was blocked when it was taller; sheet not recorded; 2026-10-02; [operator's comment on #56](https://github.com/byronbroughten/sheets-framework/issues/56#issuecomment-5962687166).
- Observed: A REST API v4 column insert at a Table's end with no `inheritFromBefore` left its new cells with no number format, background, borders or validation, beside a `NUMBER`-formatted column; the Layout tab's grid read 5 columns before and 6 after, the insert starting inside the grid; Layout tab on the dev spreadsheet, read with `exerciseLayout` dry runs either side of its send and a probe of the cells; 2026-10-09; [closing comment on #141](https://github.com/byronbroughten/sheets-framework/issues/141#issuecomment-6090354103), [recorded in 919aa4d](https://github.com/byronbroughten/sheets-framework/commit/919aa4da76040d632e65988e2de7ccd6fb4251cd).

### The grid's edge

- Observed: A REST API v4 `appendDimension` added rows or columns past the grid's edge and moved nothing; scratch tabs on the dev spreadsheet, whether a Table sat at the edge not recorded; 2026-10-03; [recorded in 7ff741d](https://github.com/byronbroughten/sheets-framework/commit/7ff741d434fda5aa91d3763c5e5da711c75ddf27).
- Observed: A REST API v4 `insertRange`, `deleteRange` or `copyPaste` whose range started past the grid's edge was refused with `Range ('<tab>'!A8:C9) exceeds grid limits. Max rows: 7, max columns: 10`, and one that ran past it was clipped to the grid; a `deleteRange` left the grid's size as it was; scratch tabs on the dev spreadsheet, whether a Table held the range not recorded; 2026-10-03; [recorded in 7ff741d](https://github.com/byronbroughten/sheets-framework/commit/7ff741d434fda5aa91d3763c5e5da711c75ddf27).
- Observed: A REST API v4 `insertRange` grew the grid only as far as the non-empty cells it pushed past the end: on a 10-row grid with values on rows 7–9, 2 rows inserted at row 5 grew it to 12; on a 514-row grid with content ending at row 503, 100 rows inserted at row 504 left it at 514, and an `updateCells` in the same batch failed with `Attempting to write row: 514, beyond the last requested row of: 513`, rejecting the batch; scratch tabs on the dev spreadsheet; 2026-10-02; [addendum on #59](https://github.com/byronbroughten/sheets-framework/issues/59#issuecomment-5963582090).

### Copying format and validation onto new rows

- Observed: A REST API v4 `copyPaste` from a Table's last body row onto rows growth added: `PASTE_FORMAT` restored the cell formatting (number format, background, text style) and grew a conditional-format rule ending at the last row from `A4:A6` to `A4:A8`, and carried no cell-level validation; `PASTE_DATA_VALIDATION` restored an untyped column's cell-level checkbox; one source row tiled over several new rows; neither touched the values or the formula fill, placed before or after them in the batch; a lone blank body row served as the source the same way; scratch tabs on the dev spreadsheet; 2026-10-02, rerun 2026-10-03; [resolution on #61](https://github.com/byronbroughten/sheets-framework/issues/61#issuecomment-5963889082), [recorded in 7ff741d](https://github.com/byronbroughten/sheets-framework/commit/7ff741d434fda5aa91d3763c5e5da711c75ddf27).
- Observed: REST API v4 cell-level `setDataValidation` on a typed Table column, `TEXT` included, was refused with `This operation is not allowed on cells in typed columns.`; scratch tab on the dev spreadsheet; 2026-10-02; [resolution on #61](https://github.com/byronbroughten/sheets-framework/issues/61#issuecomment-5963889082).

### Creating and updating Tables

- Observed: REST API v4 `addTable` wrote its own `columnProperties` names three columns right of the Table, and, on a Table whose grid ended short of those columns, replied with a 500, `Internal error encountered`; dev spreadsheet; 2026-09-23, seen again 2026-10-01; [recorded in aaa7f6a1](https://github.com/byronbroughten/byro-repo/commit/aaa7f6a12d203fe320f84c8392021a3723fcd20a), [resolution on #53](https://github.com/byronbroughten/sheets-framework/issues/53#issuecomment-5933531939).
- Observed: REST API v4 `addTable` with no `tableId` got a numeric ID in the reply, and `addTable` with `tableId: "tbl-7f3a9c21e4"` followed by an `updateTable` on that ID in the same batch set the column names, the ID reading back unchanged; scratch tab on the dev spreadsheet; 2026-10-02; [resolution on #59](https://github.com/byronbroughten/sheets-framework/issues/59#issuecomment-5963486654).
- Observed: A REST API v4 `updateTable` field mask replaced only its fields (`name`, `range`, `columnProperties`); a `columnProperties` list replaced the Table's whole list, a column left out keeping its header cell and losing its type and validation, a sent `columnName` written to the header cell, and a sent column with no name rejected; scratch tab on the dev spreadsheet; 2026-09-26, the left-out column seen again 2026-11-04; [recorded in db5ba74](https://github.com/byronbroughten/sheets-framework/commit/db5ba74d35cb24b9cf97975a1d86a589b88c7a76), [resolution on #189](https://github.com/byronbroughten/sheets-framework/issues/189#issuecomment-6624006221).
- Observed: A REST API v4 `updateTable` with `fields: "columnProperties"`, resending a Table's list as read back with each dropdown's `dataValidationRule`, turned an arrow-style dropdown column to chips and cleared a chip dropdown's option colours, both keeping their type and rule; a column rename sent the same way did the same; an `updateTable` with `fields: "name"` left style and colours as set; scratch tab on the dev spreadsheet, style and colours checked by eye; 2026-11-04; [resolution on #189](https://github.com/byronbroughten/sheets-framework/issues/189#issuecomment-6624006221).
- Observed: A REST API v4 `updateTable` `columnProperties` list giving a `DROPDOWN` column its type and no `dataValidationRule` was rejected with a 400, `Invalid requests[0].updateTable: No conditionType specified.`; scratch tab on the dev spreadsheet; 2026-11-04; [resolution on #189](https://github.com/byronbroughten/sheets-framework/issues/189#issuecomment-6624006221).
- Documented: REST API v4's `AddSheetRequest` reference says, under `properties`, "(It is an error to specify the ID of a sheet that already exists.)"; read 2026-11-01; [quoted on #196](https://github.com/byronbroughten/sheets-framework/issues/196#issuecomment-6565168668).

### The Table name chip

- Observed: In the Sheets UI, a Table's name chip sat over the cell one row above the Table's first column the whole time, and a checkbox put in that cell did not show and took no tick; a checkbox one row above the second column showed and took a tick; a Table named `aVeryLongTableNameForChipTest` showed its name truncated, over no more cells than a Table named `t`; scratch tab on the dev spreadsheet, two Tables with their header on row 4, checked by eye; 2026-11-04; [resolution on #190](https://github.com/byronbroughten/sheets-framework/issues/190#issuecomment-6633187154).

### Sorting

- Observed: A REST API v4 `sortRange` ascending put numbers, then text, then booleans (`FALSE` before `TRUE`), descending put them in the exact reverse, and blanks sorted last in both; scratch tab on the dev spreadsheet, whether a Table held the range not recorded; 2026-09-26; [recorded in db5ba74](https://github.com/byronbroughten/sheets-framework/commit/db5ba74d35cb24b9cf97975a1d86a589b88c7a76).

## Finding Tables

### Naming a Table in a data filter

- Observed: A REST API v4 `getByDataFilter` with `p130Target`, `p130Target[ID]` or `Probe130!p130Target` as its `a1Range` returned 400 "Unable to parse range", before and after each of four moves of the Table; `Probe130` tab on the dev spreadsheet, a Table; 2026-10-08; [recorded in 313ac9ba](https://github.com/byronbroughten/sheets-framework/commit/313ac9ba8a95ab532a391cb18f6f7e9509eb3255).
- Observed: A REST API v4 `getByDataFilter` with two filters, one a valid range and one an unknown name as its `a1Range`, returned 400; `Probe130` tab on the dev spreadsheet; 2026-10-08; [recorded in 313ac9ba](https://github.com/byronbroughten/sheets-framework/commit/313ac9ba8a95ab532a391cb18f6f7e9509eb3255).

### Which Tables a grid fetch returns

- Observed: A REST API v4 `getByDataFilter` grid fetch returned in `tables` only the Tables its range overlapped: a range over `layoutAbove`'s header cell alone returned `layoutAbove` only, a range over no Table returned none, a range over `p130Target`'s three head rows ending just above its header returned none, and the whole sheet as one range returned every Table on it; `Probe130` tab on the dev spreadsheet, two stacked Tables; 2026-10-08; [recorded in 313ac9ba](https://github.com/byronbroughten/sheets-framework/commit/313ac9ba8a95ab532a391cb18f6f7e9509eb3255).
- Observed: A REST API v4 `getByDataFilter` over the sheet's top rows, every column, did not return `p130Target`, its header on row 13, at 4 rows deep; at 20 rows it returned it before any move and after a row insert inside the Table above, a column insert at column A and a row delete above it, and not after a `cutPaste` of its head rows and range to H41; at 100 rows it returned it after every move; a range of column B from row 1 to the header row plus its column ID row returned it before any move and after none; `Probe130` tab on the dev spreadsheet, the moves run in turn; 2026-10-08; [recorded in 313ac9ba](https://github.com/byronbroughten/sheets-framework/commit/313ac9ba8a95ab532a391cb18f6f7e9509eb3255).
- Observed: A REST API v4 `spreadsheets.get` of the whole spreadsheet's `sheets.tables` returned `p130Target` before and after each of the four moves; dev spreadsheet; 2026-10-08; [recorded in 313ac9ba](https://github.com/byronbroughten/sheets-framework/commit/313ac9ba8a95ab532a391cb18f6f7e9509eb3255).

### Developer metadata and named ranges under moves

- Observed: REST API v4 developer metadata on `p130Target`'s header row and on its first column, looked up by `getByDataFilter`, returned the Table after a row insert inside the Table above, a column insert at column A and a row delete above it, and not after a `cutPaste` of its head rows and range to H41; the header-row lookup returned the whole sheet row, and the column lookup the whole used column with the Table above; `Probe130` tab on the dev spreadsheet; 2026-10-08; [recorded in 313ac9ba](https://github.com/byronbroughten/sheets-framework/commit/313ac9ba8a95ab532a391cb18f6f7e9509eb3255).
- Observed: A REST API v4 named range over `p130Target`'s head rows and header (B10:D13), as a `getByDataFilter` `a1Range`, returned the Table after each of the same four moves, the `cutPaste` covering the whole named range; `Probe130` tab on the dev spreadsheet; 2026-10-08; [recorded in 313ac9ba](https://github.com/byronbroughten/sheets-framework/commit/313ac9ba8a95ab532a391cb18f6f7e9509eb3255).

## Round trips and cost

### A round trip's cost

- Observed: A REST API v4 round trip cost ~250–450ms, against ~20–60ms of the framework's own CPU for a whole trigger; live spreadsheet, whether a Table held the range not recorded; 2026-09-01; [recorded in 933351eb](https://github.com/byronbroughten/byro-repo/commit/933351eb4f6cc87f6b8c82869b397f918e913cbd).
- Observed: `SpreadsheetApp`'s `getLastRow()`, `getLastColumn()` and `getRange().getValues()` together cost ~494ms, against ~350ms for the one REST API v4 `getByDataFilter` they would replace; live spreadsheet, whether a Table held the range not recorded; 2026-09-01; [recorded in 933351eb](https://github.com/byronbroughten/byro-repo/commit/933351eb4f6cc87f6b8c82869b397f918e913cbd).
- Documented: REST API v4's Usage limits page gives read requests a quota of 60 per minute per user per project, and 300 per minute per project; read 2026-11-01; [quoted on #197](https://github.com/byronbroughten/sheets-framework/issues/197#issuecomment-6566015808).

### Reads that find a Table

- Observed: REST API v4 reads from a Node script on a laptop, each timed 12 times paced under 60 reads a minute, median ms (min–max) and response bytes: column B from row 1 to the header row plus the column ID row 201 (191–228), 5,257; a named range 210 (194–241), 2,742; developer metadata on the header row 250 (225–940), 2,741, and on the first column 254 (238–1,864), 3,953; the whole spreadsheet's `sheets.tables` 188 (156–233), 13,392 over 13 sheets; the sheet's top 4 rows 202 (186–231), 2,567, top 20 rows 202 (187–230), 6,672, and top 100 rows 214 (188–250), 6,672; `Probe130` tab on the dev spreadsheet, before any move; 2026-10-08; [recorded in 313ac9ba](https://github.com/byronbroughten/sheets-framework/commit/313ac9ba8a95ab532a391cb18f6f7e9509eb3255).
- Observed: A REST API v4 `getByDataFilter` of the sheet's top 100 rows returned 8,270 bytes once a `cutPaste` had put rows 41–47 in range, against 6,672 before; `Probe130` tab on the dev spreadsheet; 2026-10-08; [recorded in 313ac9ba](https://github.com/byronbroughten/sheets-framework/commit/313ac9ba8a95ab532a391cb18f6f7e9509eb3255).
- Observed: A REST API v4 `getByDataFilter` with the whole Layout sheet as its `a1Range` and no `data` in the fields mask returned the sheet's three Tables in under 2 KB; Layout tab on the dev spreadsheet; 2026-10-08; [recorded in 313ac9ba](https://github.com/byronbroughten/sheets-framework/commit/313ac9ba8a95ab532a391cb18f6f7e9509eb3255).
- Observed: A REST API v4 grid fetch of the sheet's top 4 rows, every column, returned 1,686 bytes for both of the Layout tab's Tables, against 1,882 for two column-B strips from row 1 to each header row and their column ID rows; Layout tab on the dev spreadsheet; 2026-10-08; [recorded in cfd93d20](https://github.com/byronbroughten/sheets-framework/commit/cfd93d20e7b51e87a0687bceaa2e13fa8bb26def).

### Growing a Table in bulk

- Observed: A REST API v4 `batchUpdate` adding rows to a 20-column Table (17 value columns, 3 `SINGLE(Table[c0])` formula columns) over 500 rows, median of 3 runs each in alternating order, `appendCells` against `insertRange` plus `updateTable` plus `updateCells` plus `pasteData` formula fills: 100 rows 492ms against 533ms (76 KB against 87 KB of body), 1,500 rows 1,533ms against 1,656ms (1,174 KB against 1,337 KB), 5,000 rows 3,892ms against 4,609ms (3,956 KB against 4,499 KB), about 8% more at 1,500 rows and 18% at 5,000; an empty `GET` in the same session took 274ms; both left every new row in the Table with its formula columns filled; a fresh scratch tab per run on the dev spreadsheet, from a dev chore; 2026-10-02; [addendum on #59](https://github.com/byronbroughten/sheets-framework/issues/59#issuecomment-5963582090).

## Payload shape and reads

### Rows a grid read leaves out

- Observed: A REST API v4 `getByDataFilter` grid read of `Add Property Expenses`, whose one body row was blank and inside the Table's range (rows 3–4), returned for that row a block with 23 `columnMetadata` entries and no `rowData`, and for the column ID row a block with 23 entries and one row of 12 values, the Table being 12 columns wide; live spreadsheet, with the config sync's data filters and field mask; 2026-09-10; [measured in byro-repo#17](https://github.com/byronbroughten/byro-repo/issues/17), [recorded in bb7df180](https://github.com/byronbroughten/byro-repo/commit/bb7df18085c2737c00fca654d710ec527eff008d).

### Empty lists and zero-valued fields

- Observed: REST API v4 reads left out a sheet's `conditionalFormats` list when it held no rules, and left out zero-valued fields: a gid-0 `sheetId`, a column-A column index and a row-1 row index; live spreadsheet, the sheets and fields masks not recorded; 2026-09-16; [recorded in 62abd855](https://github.com/byronbroughten/byro-repo/commit/62abd85537ae0110cd80e3311e00c2aa298f9eb0), [recorded in 613de032](https://github.com/byronbroughten/byro-repo/commit/613de032e7a8310dd19277de344f51caca4e4ce2).

### Reading conditional formats

- Observed: A REST API v4 `getByDataFilter` returned no `conditionalFormats` for sheets holding rules, and a plain `spreadsheets.get` with the fields mask `sheets(properties(sheetId),conditionalFormats)` returned them; live spreadsheet, the sheets not recorded; 2026-09-16; [recorded in 62abd855](https://github.com/byronbroughten/byro-repo/commit/62abd85537ae0110cd80e3311e00c2aa298f9eb0), [recorded in 613de032](https://github.com/byronbroughten/byro-repo/commit/613de032e7a8310dd19277de344f51caca4e4ce2).
- Observed: A REST API v4 read of conditional-format rules returned each colour as a `*ColorStyle` (`backgroundColorStyle`, `foregroundColorStyle`) beside the plain colour field; live spreadsheet, the sheets not recorded; 2026-09-16; [recorded in 62abd855](https://github.com/byronbroughten/byro-repo/commit/62abd85537ae0110cd80e3311e00c2aa298f9eb0), [recorded in 613de032](https://github.com/byronbroughten/byro-repo/commit/613de032e7a8310dd19277de344f51caca4e4ce2).

## Formats and formulas

### Writing a formula with a Table reference

- Observed: A REST API v4 `batchUpdate` `updateCells` writing `=SUM(ProbeA[Number])` as `userEnteredValue.formulaValue` stored the text as written and it evaluated to `#ERROR!` ("Formula parse error"), in the same batch as the `addTable` and in a later one, and for `ProbeA[[Number]]`, `ProbeA[#DATA]`, `'ProbeA'[Number]` and lowercase; the same text through `values:batchUpdate` with `valueInputOption: USER_ENTERED` parsed and computed; `Table Probe` sheet on the dev spreadsheet, a Table; 2026-10-01; [resolution on #53](https://github.com/byronbroughten/sheets-framework/issues/53#issuecomment-5933531939).
- Observed: A REST API v4 read of `userEnteredValue.formulaValue` returned `=SINGLE(ProbeA[Number])*2` as written, and that text written back through `updateCells` gave a parse error; `Table Probe` sheet on the dev spreadsheet, a Table; 2026-10-01; [resolution on #53](https://github.com/byronbroughten/sheets-framework/issues/53#issuecomment-5933531939).
- Observed: A REST API v4 `repeatCell` writing a Table-reference formula as `formulaValue` left the `Formula test` cell `#ERROR!`, and a `pasteData` with `PASTE_FORMULA` of it parsed; live spreadsheet; 2026-09-14; [recorded in 3bc4a854](https://github.com/byronbroughten/byro-repo/commit/3bc4a854f6c983eefe5359d0652eb586d05dc8b4).
- Observed: In a Table's data-row formula, `SINGLE(tableName[columnName])` gave the named column's value in the formula's own row; sheet not recorded; reported 2026-11-02; [developer, quoted in the resolution on #183](https://github.com/byronbroughten/sheets-framework/issues/183#issuecomment-6595101709).
- Observed: A REST API v4 `updateTable` renaming a Table (`fields: "name"`), and one renaming a column through its `columnProperties` `columnName`, rewrote `SINGLE(table[header])` in the Table's formula column and `COUNTA(table[header])` outside the Table to the new name, their values unchanged; scratch tab on the dev spreadsheet; 2026-11-04; [resolution on #189](https://github.com/byronbroughten/sheets-framework/issues/189#issuecomment-6624006221).

### Number formats on blank cells

- Observed: A REST API v4 grid read returned Occupancy's `Next terms notes` first body cell blank, with no declared column type and a `CURRENCY` number format; live spreadsheet, a Table; 2026-09-10; [recorded in 35858b5b](https://github.com/byronbroughten/byro-repo/commit/35858b5b081de15c2b415a0dc89d9d67fd7126a1).

### Checkbox cells

- Observed: `Add Property Expense`'s one body row held an explicit `FALSE` in a just-added checkbox column and nothing anywhere else; live spreadsheet, a Table, how the column was added and the read's surface not recorded; 2026-09-11; [recorded in 4e93f46f](https://github.com/byronbroughten/byro-repo/commit/4e93f46f392780ef5fc25ecd87364776269f4ef2).

## Protections

### Lock editors

- Observed: A lock added through REST API v4 naming no editors read back with the spreadsheet's editors (two gmail users and a service account), and in the Sheets UI the owner edited its cells (Num) with no prompt and no block; Table data column on `test`; 2026-09-17; [recorded in a03c0a6a](https://github.com/byronbroughten/byro-repo/commit/a03c0a6af3e6ba86a5eafcbc229d4a90333d06af).
- Observed: A lock added through REST API v4 naming users read back with exactly those users, none of the spreadsheet's other editors added; `test`; 2026-09-17; [recorded in a03c0a6a](https://github.com/byronbroughten/byro-repo/commit/a03c0a6af3e6ba86a5eafcbc229d4a90333d06af).
- Observed: With a lock added through REST API v4 naming one gmail address on Column currency's bounded data range, the Sheets UI blocked the spreadsheet owner's edits to those cells, and an edit to the last row, which sat outside the range, went through; the address sent and the address read back were not kept; `test`; 2026-09-17; [recorded in a03c0a6a, corrected in 28fbbd79](https://github.com/byronbroughten/byro-repo/commit/28fbbd792bea56805b6edc54646280ef60fe50de).
- Observed: The Sheets UI refused to create a lock that would lock the acting user out, and REST API v4 accepted such a lock; the sheet was not recorded; 2026-09-17; [recorded in 28fbbd79](https://github.com/byronbroughten/byro-repo/commit/28fbbd792bea56805b6edc54646280ef60fe50de).

### Reading protections back

- Observed: A REST API v4 `getByDataFilter` returned the same `protectedRanges` payload as a plain get with a `protectedRanges` fields mask; `test`; 2026-09-17; [recorded in a03c0a6a](https://github.com/byronbroughten/byro-repo/commit/a03c0a6af3e6ba86a5eafcbc229d4a90333d06af).
- Observed: An open-ended warning added through REST API v4 (start row and column bounds, no `endRowIndex`) read back with no `endRowIndex`, as `G5:G`; Table data column on `test`; 2026-09-17; [recorded in a03c0a6a](https://github.com/byronbroughten/byro-repo/commit/a03c0a6af3e6ba86a5eafcbc229d4a90333d06af).
- Observed: A warning added through REST API v4 with no editors read back with an `editors` list; `test`; 2026-09-17; [recorded in a31f1523](https://github.com/byronbroughten/byro-repo/commit/a31f1523b33625f3fe9ca136ac995703f232aa71).
- Observed: After the probe's four protections were deleted, a REST API v4 get with a `protectedRanges` fields mask returned `test` with no `protectedRanges` field; 2026-09-17; [recorded in a03c0a6a](https://github.com/byronbroughten/byro-repo/commit/a03c0a6af3e6ba86a5eafcbc229d4a90333d06af).
- Documented: REST API v4's `ProtectedRange.range` reference says a fully unbounded range is considered a protected sheet, and its `GridRange` reference says a missing index leaves the range unbounded on that side; read 2026-10-31; [quoted on #195](https://github.com/byronbroughten/sheets-framework/issues/195#issuecomment-6553553665).
- Documented: REST API v4's `ProtectedRange` reference says that on a write with `warningOnly` true the `editors` are ignored, and that editors are not supported with `warningOnly` protection; read 2026-10-31; [quoted on #195](https://github.com/byronbroughten/sheets-framework/issues/195#issuecomment-6553553665).

### Rows added below a protection

- Observed: Bounded data-column protections added through REST API v4 (`B5:B16`, `E5:E16`, `I5:I16`) kept their end row through an `appendCells` that added Table row 17 and through typing below the Table in the Sheets UI, and the appended row was unprotected in the UI; Table on `test`; 2026-09-17; [recorded in a03c0a6a](https://github.com/byronbroughten/byro-repo/commit/a03c0a6af3e6ba86a5eafcbc229d4a90333d06af).
- Observed: An open-ended warning added through REST API v4 (no `endRowIndex`) reaching below the data prompted in the Sheets UI on a cell under the Table, and on a cell in a Table row added afterwards; Table on `test`; 2026-09-17; [recorded in 28fbbd79](https://github.com/byronbroughten/byro-repo/commit/28fbbd792bea56805b6edc54646280ef60fe50de).
- Observed: A protection with a concrete end row kept that end row when 1000 rows were added to the sheet; `test`; 2026-09-17; [recorded in 28fbbd79](https://github.com/byronbroughten/byro-repo/commit/28fbbd792bea56805b6edc54646280ef60fe50de).
- Observed: `$E5:E` typed into the Sheets UI's protected-range editor was saved with a concrete end row, the sheet's last row at the time; `test`; 2026-09-17; [recorded in 28fbbd79](https://github.com/byronbroughten/byro-repo/commit/28fbbd792bea56805b6edc54646280ef60fe50de).

### Warning prompts

- Observed: A warning added through REST API v4 on the ID column prompted the spreadsheet owner on an edit in the Sheets UI; Table data column on `test`; 2026-09-17; [recorded in a03c0a6a](https://github.com/byronbroughten/byro-repo/commit/a03c0a6af3e6ba86a5eafcbc229d4a90333d06af).
- Observed: A cell carrying both a lock that permitted the editing user and a warning prompted on an edit in the Sheets UI; `test`; 2026-09-17; [recorded in 28fbbd79](https://github.com/byronbroughten/byro-repo/commit/28fbbd792bea56805b6edc54646280ef60fe50de).
- Observed: A warning's `description` showed in the Sheets UI's Protected sheets and ranges sidebar and not in its prompt, which carried no custom message; `test`; 2026-09-17; [recorded in 28fbbd79](https://github.com/byronbroughten/byro-repo/commit/28fbbd792bea56805b6edc54646280ef60fe50de).
- Observed: With the probe's warnings and locks on its data columns, the Sheets UI refused a column drag for every column; `test`; 2026-09-17; [recorded in a03c0a6a](https://github.com/byronbroughten/byro-repo/commit/a03c0a6af3e6ba86a5eafcbc229d4a90333d06af).
- Documented: REST API v4's `ProtectedRange.warningOnly` reference says every user edits data in a warning-based protected range, with a warning asking the user to confirm the edit; read 2026-10-31; [quoted on #195](https://github.com/byronbroughten/sheets-framework/issues/195#issuecomment-6553553665).

### Whole-sheet warnings and unprotected ranges

- Observed: A whole-sheet warning prompted in the Sheets UI on a tab rename, a tab delete and a column insert; sheet and how the warning was added not recorded; 2026-09-18; [recorded in 319e5673](https://github.com/byronbroughten/byro-repo/commit/319e5673bd29b9546b1f75c8246ef42d2f9c4b15).
- Observed: An edit in the Sheets UI inside one of a whole-sheet warning's unprotected ranges did not prompt; sheet and how the warning was added not recorded; 2026-09-18; [recorded in 319e5673](https://github.com/byronbroughten/byro-repo/commit/319e5673bd29b9546b1f75c8246ef42d2f9c4b15).
- Observed: A column inserted strictly inside a whole-sheet warning's unprotected range was inside that range afterwards, and one inserted at either of its edges was outside it; sheet, how the warning was added and the inserts' surface not recorded; 2026-09-18; [recorded in 319e5673](https://github.com/byronbroughten/byro-repo/commit/319e5673bd29b9546b1f75c8246ef42d2f9c4b15).
- Observed: On a whole-sheet warning, a column insert in the Sheets UI inside a whole-grid unprotected range prompted, and a column insert strictly inside a whole-column unprotected range did not; sheet and how the warning was added not recorded; 2026-09-18; [recorded in 319e5673](https://github.com/byronbroughten/byro-repo/commit/319e5673bd29b9546b1f75c8246ef42d2f9c4b15).

### Row and column shifts

- Observed: A protection's unprotected range of `B6:D9` read back as `B5:D8` after a row above it was deleted; live spreadsheet, with the sheet, the delete's surface and whether a Table held the range not recorded; 2026-09-19; [recorded in c3dedd7a](https://github.com/byronbroughten/byro-repo/commit/c3dedd7ab940dae089feb7ab688fc482b55d850f).
- Observed: REST API v4 row and column inserts and deletes moved the sheet's protected ranges with the rows and columns; scratch tab on the dev spreadsheet; 2026-09-26; [recorded in db5ba74d](https://github.com/byronbroughten/sheets-framework/commit/db5ba74d35cb24b9cf97975a1d86a589b88c7a76).

### Adding protections

- Documented: REST API v4's `AddProtectedRangeResponse` reference gives the reply one field, `protectedRange`, "the newly added protected range", a `ProtectedRange`, whose read-only `protectedRangeId` is the range's id; read 2026-10-31; [quoted on #195](https://github.com/byronbroughten/sheets-framework/issues/195#issuecomment-6553553665).

### Deleting protections

- Observed: A REST API v4 `deleteProtectedRange` with an id no sheet held was rejected with `No protected range with id`; scratch tab on the dev spreadsheet; 2026-09-26; [recorded in 19cc66e2](https://github.com/byronbroughten/sheets-framework/commit/19cc66e2e3f7de98a3c0edb2a66a6e855e9d0f8c).

## Apps Script services

### The advanced Sheets service

- Documented: Apps Script's advanced Sheets service page says the service "uses the same objects, methods, and parameters as the public API"; read 2026-10-01; [write-up on research/google-sheets-tables](https://github.com/byronbroughten/sheets-framework/blob/research/google-sheets-tables/research/google-sheets-tables.md#6-apps-script-visibility).

## Triggers

### Simple triggers and authorization

- Documented: Apps Script's Simple Triggers guide says, under its restrictions, that simple triggers "can't access services that require authorization"; read 2026-11-02; [quoted on #199](https://github.com/byronbroughten/sheets-framework/issues/199#issuecomment-6590878536).

### Toasts from an installed edit trigger

- Observed: `SpreadsheetApp.getActive().toast(...)`, called from an installable on-edit trigger that account A installed, showed to A when A ticked a checkbox, and to account B, another editor, when B ticked; A, with the spreadsheet open in another browser profile, saw no toast for B's tick; scratch tab on the dev spreadsheet; 2026-11-04; [resolution on #190](https://github.com/byronbroughten/sheets-framework/issues/190#issuecomment-6633187154).

### Edit event ranges

- Documented: `SpreadsheetApp`'s `Range.getColumn()` reference says it "Returns the starting column position for this range", with an example on `B2:D4` commented `// Logs "2.0"`; read 2026-11-02; [quoted on #199](https://github.com/byronbroughten/sheets-framework/issues/199#issuecomment-6590878536).

## Auth

### Refresh tokens under a Testing consent screen

- Documented: Google Identity's Using OAuth 2.0 page says a Google Cloud project whose OAuth consent screen is set to the external user type with a publishing status of "Testing" "is issued a refresh token expiring in 7 days", unless the only scopes requested are name, email address and user profile; read 2026-09-27, again 2026-11-02; [write-up on research/harness-auth](https://github.com/byronbroughten/sheets-framework/blob/research/harness-auth/research/harness-auth.md#scopes-and-verification), [quoted on #199](https://github.com/byronbroughten/sheets-framework/issues/199#issuecomment-6590878536).
