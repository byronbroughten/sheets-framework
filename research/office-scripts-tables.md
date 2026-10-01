# How Office Scripts and the Excel JavaScript API model tables

Research for byronbroughten/sheets-framework#45. Sources were read on 2026-10-01. Microsoft Learn API reference and Microsoft Support articles are primary sources. Anything that comes only from forum answers or from my own knowledge is marked **[secondary]** or **[unverified]**.

Two APIs share one object model:

- **Excel JavaScript API** (`Excel.*`, used by add-ins): proxy objects, `load()` + `context.sync()`, collections such as `TableCollection` and `TableRowCollection`.
- **Office Scripts** (`ExcelScript.*`): a synchronous flavour of the same API. It uses `getX()`/`setX()` methods and plain arrays instead of collection objects. It has **no `TableRow` object**.

## 1. Workbook-level vs worksheet-level access

- In Excel JS, the same `TableCollection` class hangs off both `Workbook.tables` and `Worksheet.tables`. It "Represents a collection of all the tables that are part of the workbook or worksheet, depending on how it was reached." ([Excel.TableCollection](https://learn.microsoft.com/en-us/javascript/api/excel/excel.tablecollection))
- `TableCollection.getItem(key)` "Gets a table by name or ID". `getItemOrNullObject(key)` is the non-throwing variant. `getItemAt(index)` takes a position. ([Excel.TableCollection](https://learn.microsoft.com/en-us/javascript/api/excel/excel.tablecollection))
- In Office Scripts, `Workbook.getTable(key)` "Gets a table by name or ID. If the table doesn't exist, then this method returns `undefined`", and `Workbook.getTables()` returns the workbook's tables. ([ExcelScript.Workbook](https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.workbook))
- `Worksheet.getTable(key)` behaves the same way, and `Worksheet.getTables()` is the "Collection of tables that are part of the worksheet." ([ExcelScript.Worksheet](https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.worksheet))
- You can also reach tables from a range. `Range.getTables(fullyContained?)` "Gets a scoped collection of tables that overlap with the range." ([Excel.Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range)) The Office Scripts sample relies on this: "Since tables can't overlap, this will be one table at most." ([Table samples](https://learn.microsoft.com/en-us/office/dev/scripts/resources/samples/table-samples))
- **Names are workbook-scoped.** Table names must be unique ("Duplicate names aren't allowed. Excel doesn't distinguish between upper and lowercase characters in names"). A name must begin with a letter, `_` or `\`, contain no spaces, not look like a cell reference such as `Z$100` or `R1C1`, and be at most 255 characters long. ([Rename an Excel table](https://support.microsoft.com/office/fbf49a4f-82a3-43eb-8ba2-44d21233b114), linked from [Excel.Table.name](https://learn.microsoft.com/en-us/javascript/api/excel/excel.table))
  - This is why a workbook-level `getTable(name)` lookup is unambiguous, and why the worksheet-level lookup is only a filter.
- **Every table has a stable ID that survives renames.** `Table.id` "uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed." ([Excel.Table](https://learn.microsoft.com/en-us/javascript/api/excel/excel.table); [ExcelScript.Table.getId](https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.table)) There is also a numeric `legacyId`.
- **Each table knows its sheet.** `Table.worksheet` / `getWorksheet()` returns "The worksheet containing the current table." ([Excel.Table](https://learn.microsoft.com/en-us/javascript/api/excel/excel.table))

## 2. How many tables per worksheet, and the constraints

- **Count:** none of the API or support pages I read caps the number of tables on a worksheet. Multiple tables per sheet is normal: the docs show `sheet.tables.add(...)` repeatedly, and `Worksheet.getTables()` returns an array. **[unverified]** I found no explicit limit in the sources I read.
- **No overlap.** `add(address, hasHeaders)` says: "If the table cannot be added (e.g., because the address is invalid, or the table would overlap with another table), an error will be thrown." ([Excel.TableCollection](https://learn.microsoft.com/en-us/javascript/api/excel/excel.tablecollection); same wording on [ExcelScript.Worksheet.addTable](https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.worksheet))
  - The `MergedRangeConflict` error code is broader: "A table can't overlap with another table, a PivotTable report, query results, merged cells, or an XML Map." ([Excel error codes](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-error-handling))
- **Adjacency:** no doc forbids tables that touch. In practice, adjacency is what makes row and column shifts fail (§6).
  - Excel also auto-expands a table when you type directly below or to the right of it: "the table automatically expands to include that cell." ([Resize a table](https://support.microsoft.com/en-us/office/resize-a-table-by-adding-or-removing-rows-and-columns-e65ae4bb-e44b-43f4-ad8b-7d68784f1165)) So a table pressed against unrelated data or another table invites surprises.
- **Header row requirement:**
  - On creation, `hasHeaders: false` means "Excel will automatically generate a header and shift the data down by one row." ([Excel.TableCollection.add](https://learn.microsoft.com/en-us/javascript/api/excel/excel.tablecollection)) So a table always has column names internally.
  - Afterwards the header row can be hidden: `showHeaders` "can be set to show or remove the header row." ([Excel.Table](https://learn.microsoft.com/en-us/javascript/api/excel/excel.table))
  - Filter buttons need a header: `showFilterButton` "Setting this is only allowed if the table contains a header row." (same page)
  - Column names are the table's schema. Structured references and `getColumnByName` depend on them.
- **Resize rule:** "The new range must overlap with the original table range and the headers (or the top of the table) must be in the same row." ([Excel.Table.resize](https://learn.microsoft.com/en-us/javascript/api/excel/excel.table)) Support states the same: "Table headers can't move to a different row, and the new range must overlap the original range." ([Resize a table](https://support.microsoft.com/en-us/office/resize-a-table-by-adding-or-removing-rows-and-columns-e65ae4bb-e44b-43f4-ad8b-7d68784f1165))
  - This makes the header row the table's anchor.

## 3. The parts of a table and their ranges

Each part of a table is exposed as a method that returns a plain `Range`. The parts are not separate object types.

| Part | Excel JS | Office Scripts |
| --- | --- | --- |
| Whole table | `getRange()` | `getRange()` |
| Header row | `getHeaderRowRange()` | `getHeaderRowRange()` |
| Data body | `getDataBodyRange()` | `getRangeBetweenHeaderAndTotal()` |
| Total row | `getTotalRowRange()`, toggled by `showTotals` | `getTotalRowRange()`, `setShowTotals()` |

Sources: [Excel.Table](https://learn.microsoft.com/en-us/javascript/api/excel/excel.table), [ExcelScript.Table](https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.table).

**Columns** are first-class objects with identity:

- `TableColumn` has a `name`, a zero-based `index` and an `id`. "The `id` property of a `TableColumn` object contains a unique key that identifies the column." ([Excel tables how-to](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-tables)) Office Scripts: `getId()` "Returns a unique key that identifies the column within the table." ([ExcelScript.TableColumn](https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.tablecolumn))
- Lookup methods:
  - `getColumn(key: number | string)` takes a name or ID.
  - `getColumnById` and `getColumnByName` return `undefined` when the column is missing.
  - `getColumns()` returns all columns.
  - `addColumn(index, values, name)` adds one.
  - Source: [ExcelScript.Table](https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.table).
  - **Pitfall:** the Office Scripts sample calls `table.getColumn(2)` with a comment saying "1-based indices". A number passed there is treated as a column **ID**, not a position. IDs happen to start at 1 in a fresh table. ([Table samples](https://learn.microsoft.com/en-us/office/dev/scripts/resources/samples/table-samples))
- Each column exposes its own slices: `getHeaderRowRange()`, `getRangeBetweenHeaderAndTotal()` (Excel JS: `getDataBodyRange()`), `getTotalRowRange()` and `getRange()`. ([ExcelScript.TableColumn](https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.tablecolumn))
- Filters hang off the column (`getFilter()`). Sort hangs off the table (`getSort()`). ([ExcelScript.Table](https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.table))

## 4. How a row is addressed

**Rows are positional and have no identity.**

- "A `TableRow` object doesn't have an `id` property, so use the row position when you need to identify a row." ([Excel tables how-to](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-tables))
- Docs for `TableRow`, `TableRowCollection`, `getItemAt` and `add` all repeat the same note: "unlike ranges or columns, which will adjust if new rows or columns are added before them, a `TableRow` object represents the physical location of the table row, but not the data. That is, if the data is sorted or if new rows are added, a table row will continue to point at the index for which it was created." ([Excel.TableRow](https://learn.microsoft.com/en-us/javascript/api/excel/excel.tablerow), [Excel.TableRowCollection](https://learn.microsoft.com/en-us/javascript/api/excel/excel.tablerowcollection))

**Excel JS API (`TableRowCollection`, at `table.rows`)**:

- `add(index?, values?, alwaysInsert?)`:
  - `index` null or -1 means the end. "Any rows below the inserted row are shifted downwards. Zero-indexed."
  - `alwaysInsert` (default `true`): "If `true`, the new rows will be inserted into the table. If `false`, the new rows will be added below the table."
  - It returns the top new `TableRow`.
- `getItemAt(index)`, `getCount()` / `count`, `items`.
- `deleteRows(rows: number[] | TableRow[])` deletes rows that "don't need to be sequential". `deleteRowsAt(index, count?)` deletes a run of rows.
  - Both throw `InsertDeleteConflict` "if the table on which the method is called has a filter applied."
  - On indexes: "Row indexes update each time that a preceding row in the table is deleted."
  - The delete methods are in `ExcelApiOnline 1.1`.
- Source for all of the above: [Excel.TableRowCollection](https://learn.microsoft.com/en-us/javascript/api/excel/excel.tablerowcollection).
- `TableRow` has `index`, `values` (2D), `valuesAsJson`, `getRange()` and `delete()`. ([Excel.TableRow](https://learn.microsoft.com/en-us/javascript/api/excel/excel.tablerow))

**Office Scripts** has no row object. Rows are handled through methods on `Table`:

- `addRow(index?, values?)` takes a 1-D array. `addRows(index?, values?)` takes a 2-D array.
- `deleteRowsAt(index, count?)`, documented with these cautions:
  - "Caution: the index of the row may have moved from the time you determined the value to use for removal."
  - "Deleting more than 1000 rows at the same time could result in a Power Automate timeout."
- `getRowCount()`.
- Source: [ExcelScript.Table](https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.table).
- To read a row, you slice `getRangeBetweenHeaderAndTotal()` and use `getValues()`.

## 5. Structured references in formulas

- **Definition:** "That combination of table and column names is called a structured reference. The names in structured references adjust whenever you add or remove data from the table." ([Using structured references](https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e))
- **Syntax:**
  - A table name, such as `DeptSales`, "references the table data, without any header or total rows".
  - Column specifiers use brackets: `[Sales Amount]`. A column range uses `DeptSales[[Sales Person]:[Region]]`.
  - Item specifiers:
    - `#All`: "The entire table, including column headers, data, and totals (if any)."
    - `#Data`: "Just the data rows."
    - `#Headers`: "Just the header row."
    - `#Totals`: "Just the total row. If none exists, then it returns null."
    - `#This Row` / `@`: "Just the cells in the same row as the formula."
  - "All table, column, and special item specifiers need to be enclosed in matching brackets." Special characters in a header need doubled brackets, and some need a `'` escape.
  - Source for this list: [Using structured references](https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e).
- **Qualified vs unqualified:** inside the table (a calculated column), unqualified `[@Col]` is allowed. Outside the table "you need to use a fully qualified structured reference." (same page)
- **Renames:** "If you rename a column or table, Excel automatically changes the use of that table and column header in all structured references." (same page)
  - This is the durability property that A1 references lack.
  - The Office Scripts sample states the scripting payoff: "The column names in the table can be changed without changing this script", using `` `=[@[${name2}]]*[@[${name3}]]` `` written into `getRangeBetweenHeaderAndTotal()`. ([Table samples](https://learn.microsoft.com/en-us/office/dev/scripts/resources/samples/table-samples))
- **Calculated columns:** writing one formula into a column body makes Excel "automatically create a calculated column and copies the formula down the entire column." ([Using structured references](https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e); [Overview of Excel tables](https://support.microsoft.com/en-us/office/overview-of-excel-tables-7ab0bb7d-3a9e-4b56-a3c9-6c94334e492c))
- **Copying vs filling:** "When you copy, all the structured references remain the same, while when you fill a formula, fully qualified structured references adjust the column specifiers like a series." ([Using structured references](https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e))
- **Converting to a range:** "all cell references change to their equivalent absolute A1 style references." (same page)
- **Column moves and deletes:** **[unverified]** The support page does not say this explicitly. Because references are name-based, moving a column within a table keeps formulas pointing at that column. Deleting a referenced column turns references to it into `#REF!`.
- **External references:** a workbook linking to a table in another workbook needs that source workbook open "to avoid #REF! errors." ([Using structured references](https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e))

## 6. Worksheet row insert/delete when several tables share a sheet

**What the primary docs say:**

- The API surfaces these failures as error codes rather than as messages:
  - `InsertDeleteConflict`: "The insert or delete operation attempted resulted in a conflict."
  - `FilteredRangeConflict`: "The attempted operation causes a conflict with a filtered range."
  - `NonBlankCellOffSheet`: "can't insert new cells because it would push non-empty cells off the end of the worksheet."
  - Source: [Excel error codes](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-error-handling).
- Table row deletes refuse to run on a filtered table and throw `InsertDeleteConflict`. ([Excel.TableRowCollection](https://learn.microsoft.com/en-us/javascript/api/excel/excel.tablerowcollection))
- `Range.insert(shift)` "shifts the other cells to make space". `Range.delete(shift)` takes `"Up" | "Left"`. Neither reference page documents table-specific behaviour. ([Excel.Range](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range))
- Excel UI distinguishes table operations from sheet operations: "Table Rows Above/Below" and "Delete Table Rows" are separate from sheet row deletion. ([Resize a table](https://support.microsoft.com/en-us/office/resize-a-table-by-adding-or-removing-rows-and-columns-e65ae4bb-e44b-43f4-ad8b-7d68784f1165))

**The UI message.** "This operation is not allowed. The operation is attempting to shift cells in a table on your worksheet."

- **[secondary]** A Microsoft Support agent on Microsoft Q&A explains: "Normally, you can use Insert to create a new row in a sheet, but if that table will then overlap with an existing table, the Insert command will throw the error." A moderator adds: "If more than one [table], make sure you are selecting only the row in one table." Stray values near the bottom of the sheet also block inserts. ([Q&A: Excel not allowing new rows](https://learn.microsoft.com/en-us/answers/questions/5108506/excel-not-allowing-new-rows))
- **[secondary]** The same message appeared in VSTO only after "another list object [was added] to the same sheet." ([MSDN forum archive](https://learn.microsoft.com/en-us/archive/msdn-technet-forums/17879c52-8214-4af4-ba77-b70fa53c1612))

**How it works out in practice.** **[unverified]** This is synthesis from the above plus general Excel knowledge, not from a Microsoft page:

- **Entire-row insert or delete on the sheet:**
  - Tables stacked vertically just shift together, so this usually succeeds.
  - Tables side by side that the row cuts through also usually succeed: each table grows or shrinks by the row.
  - It fails when the result would make tables overlap, would cut a table's header out, or would push data off the sheet.
- **Partial-row (cell-range) insert or delete with shift Down/Up:** this is the common failure.
  - The shift moves cells in one column band only.
  - If that band crosses part of another table below, or partly covers a table's columns, Excel refuses with the "shift cells in a table" message.
  - This is why `ListRows.Add` / `rows.add` on a table that has another table directly beneath it can fail.
- **Table-scoped insert/delete** (`rows.add`, `deleteRowsAt`): this shifts cells only within the table's columns, below the target row.
  - It succeeds when those columns below the table are free.
  - It fails when another table or a filter is in the way.
  - Office Scripts `addRow`/`addRows` behave the same way.
- **Net effect:** Excel treats a table as a rigid rectangle. Any shift that would split, overlap or partially move a table is rejected rather than silently corrupting data.
  - Layouts with several tables per sheet stay reliable when tables are separated by blank columns or rows and have nothing stacked directly below them.

## 7. What a table-centric framework over Google Sheets should copy or avoid

Context: the framework currently has Spreadsheet, Sheet, SheetMeta, Column, ColumnMeta, Row, UniformRow and Cell, and is considering `spreadsheet.table(name)`, `table.header(i)` and `table.row(i)`.

**Copy:**

1. **Look tables up by name at workbook level.** Excel makes table names unique across the workbook, so `spreadsheet.table(name)` needs no sheet qualifier. Keep `sheet.tables()` as a filter. Every table should also expose its sheet (`table.sheet`), as Excel's `Table.worksheet` does.
2. **Give tables a stable ID alongside the name.** Excel's `Table.id` survives renames. Key generated configs and caches by that ID; treat the name as a mutable label.
3. **Make columns first-class with name, position and ID.** Match Excel's `getColumnByName` / `getColumnById` / position triad. Don't overload one numeric key the way Office Scripts `getColumn(2)` does, where the number is an ID and samples mistake it for a position. Make `header(i)` unambiguously positional, and add `column(name)`.
4. **Expose the parts of a table as named slices**: header, body and totals, both per table and per column. Excel's `getHeaderRowRange` / `getDataBodyRange` / `getTotalRowRange` returning ordinary ranges is a small, deep interface. A `body` slice matches the existing Row/UniformRow split well.
5. **Make the header row the anchor and keep it fixed.** Excel's resize rule pins headers to one row and requires overlap. A framework can enforce the same thing: a table's header row never moves during a resize, only the body grows.
6. **Fail loudly instead of corrupting data.** Excel refuses shifts that would split or overlap tables (`InsertDeleteConflict`, `MergedRangeConflict`). Check that an insert or delete stays within its table, and that the space below is free, before queuing writes, and give the failure a named error.
7. **Use name-based formula references.** Structured references survive column and table renames. Google Sheets now has its own table references, which are worth a separate check. If they aren't usable, generating formulas from column names at write time, as the Office Scripts sample does, gets similar durability.

**Avoid:**

1. **Rows as durable handles.** Excel says plainly that a `TableRow` "represents the physical location … but not the data" and goes stale after a sort or insert. If `table.row(i)` returns an object, scope it to one read/write pass, or key rows by a domain ID column instead of by index. Never cache row objects across mutations.
2. **Deleting by index without re-reading.** Both APIs warn that indexes move between lookup and delete. Delete in descending index order, or delete a batch by identity in one call, as `deleteRows([...])` does.
3. **Treating sheet-level row operations as table operations.** Excel keeps "Delete Table Rows" separate from sheet row deletion. A table-centric API should offer table-scoped insert and delete. Sheet-wide inserts should be explicit, because they affect every table the row crosses.
4. **Stacking tables directly on top of each other.** In Excel this is what makes `rows.add` and shifts fail, and auto-expansion swallows adjacent data. If one sheet holds several tables, recommend or enforce side-by-side layout with gutters, or one growable table per column band.
5. **Mutating a filtered view.** Excel refuses row deletes on filtered tables. Google Sheets filter views also hide rows, so blank-row and deletion paths should refuse to run while a filter is active, or clear it first.
