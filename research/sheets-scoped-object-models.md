# Which Sheets libraries offer a scoped object model, and how cheap is each for an agent?

Research for [#34](https://github.com/byronbroughten/sheets-framework/issues/34) (map [#24](https://github.com/byronbroughten/sheets-framework/issues/24)). It builds on [sheets-tooling-landscape.md](https://github.com/byronbroughten/sheets-framework/blob/research/sheets-tooling-landscape/research/sheets-tooling-landscape.md) (#25) and doesn't repeat it. Gathered 2026-09-30 from official reference docs, shallow clones of each library's default branch (commit noted per library), the published npm packages and the GitHub API. Two Excel libraries, Office.js's Excel JavaScript API and ExcelJS, are covered for inspiration only, in [their own section](#excel-for-inspiration).

"Table" below means a native Sheets table (`addTable`/`updateTable`/`deleteTable`, `tableId`), unless a library's own class is called out as a header-row abstraction.

## Summary

- **No Sheets library has a native Table scope.** Not Apps Script `SpreadsheetApp`, not gspread, not google-spreadsheet, not pygsheets, sheetfu, ezsheets or google-drive-ruby. The only native-table helpers in the whole ecosystem are the curated MCP tools that #25 found (taylorwilsdon, a-bonus), and those are tools, not a class model.
- **Only a minority batch implicitly.** gspread, ezsheets and most google-spreadsheet helpers send one HTTP call per operation, and batching means calling a separate `batch_*` method. sheetfu (`batch_to=`), pygsheets (`set_batch_mode`) and google-drive-ruby (`save`) can queue scoped operations into one `batchUpdate`. Apps Script bundles opaquely inside the runtime.
- **A raw escape hatch is nearly universal** but always untyped beside the helpers: `Spreadsheet.batch_update(dict)` in gspread, a `ky` instance in google-spreadsheet, the Advanced Sheets service in Apps Script.
- **Agent cost splits in two.** Apps Script is the only model with a typed surface, and it is large: 68 interfaces and 1,404 method signatures in `@types/google-apps-script`'s spreadsheet file. The small libraries (sheetfu, ezsheets) are cheap but untyped Python and unmaintained. google-spreadsheet is the only small, typed, maintained one, and it has no tables and no general batching.

- **Excel has already solved the design, but not for Sheets.** Office.js has exactly the target shape: Workbook → Worksheet → Table → TableColumn/TableRow → Range proxies that queue commands until one `context.sync()`. Its table scope is small: `Table`, `TableColumn` and `TableRow` together have 31 methods. ExcelJS has a file-level `Table` whose edits are cached and applied by `commit()`. Neither can reach a Google spreadsheet.

So "parsimonious scoped classes, table included, compiling to one `batchUpdate`, typed, with raw requests alongside, cheap for an agent" is **open ground for Sheets**. Each property exists somewhere, but no Sheets library combines more than two of them. Office.js is the design to borrow from.

## Comparison

| Library | Lang | Scopes | Native Table scope | Batching of scoped ops | Raw request beside helpers | Agent-cost measure | Health (2026-09-30) |
| --- | --- | --- | --- | --- | --- | --- | --- |
| **Apps Script `SpreadsheetApp`** | JS (V8, in Google's runtime) | SpreadsheetApp, Spreadsheet, Sheet, Range, RangeList; no Row/Column/Cell classes (a Range of one row or cell) | **No** | Implicit and opaque: "Spreadsheet operations are sometimes bundled together to improve performance" (`flush()`) | Yes: the Advanced Sheets service calls the REST API, untyped from the helpers' side | Typed and large: 68 interfaces, 1,404 method signatures, 2,607 lines; Range 245, Sheet 157, Spreadsheet 133 | First-party. No Spreadsheet-service release notes in 2025-2026 |
| **gspread** | Python | Client, Spreadsheet, Worksheet, Cell; rows and columns are methods | No | One call each. Explicit batch methods: `Worksheet.batch_update` (values), `batch_format`, `batch_clear`, `batch_merge` | Yes: `Spreadsheet.batch_update(body: Mapping)`, untyped | 183 public methods (Worksheet 100, Spreadsheet 41, Client 16, Cell 5, HTTPClient 21); 7.4k lines; typed (`py.typed`) but dict-shaped | 7.5k stars; v6.2.1 (2025-05-14); last push 2026-07-30 |
| **google-spreadsheet** (npm, theoephraim) | TypeScript | GoogleSpreadsheet, Worksheet, Row, Cell | No | Mixed. Cells batch: `saveUpdatedCells()` sends every dirty cell in one `batchUpdate`. Rows don't: `row.save()` is one `values` PUT and `row.delete()` one `batchUpdate` each | Yes: public `sheetsApi` `ky` instance, documented for "unsupported sheets calls"; untyped | 5 classes, 1,571-line `index.d.ts`, 3.0k lines of source; 11k words of docs | 2.5k stars; v5.3.0; last push 2026-09-04 |
| **pygsheets** | Python | Client, Spreadsheet, Worksheet, DataRange, Cell, Address/GridRange | No | Opt-in client-wide: `set_batch_mode(True)` queues `batchUpdate` requests, not value updates or clears; `run_batch()` sends them | Yes: `client.sheet.batch_update(spreadsheet_id, requests)` | 222 public methods across the six scope classes; 6.4k lines | 1.5k stars; v2.0.5; last push 2025-06-10 |
| **sheetfu** | Python | Spreadsheet, Sheet, Range, plus a `Table`/`Item` module | No. Its `Table` is a header-row range with dict items | Per call: every setter takes `batch_to=` (a Spreadsheet, Sheet or Table); `commit()` sends one `batchUpdate`. Mirrors Apps Script's naming | Yes: `client.sheet_service` is the raw `googleapiclient` service | 63 public methods (model 40, table 23); 1.7k lines | 846 stars; last commit 2022-03-16 |
| **ezsheets** | Python | Spreadsheet, Sheet; `getRow`/`updateColumn` etc. on Sheet | No | One call each | No public passthrough | One 1.9k-line module; aimed at beginners | 75 stars; v2024.8.9 |
| **google-drive-ruby** | Ruby | Session, Spreadsheet, Worksheet; `ws[r, c] = v` cell access | No | Buffered: `[]=` edits queue `@v4_requests`; `save` sends one `batchUpdate` | Yes: public `Spreadsheet#batch_update(requests)` | Small Worksheet API | 1.8k stars; last commit 2021-04-13 |
| *Office.js Excel API (inspiration)* | TypeScript (add-in runtime) | Workbook, Worksheet, **Table**, TableColumn, TableRow, Range (no Cell class) | **Yes**, native Excel tables | All proxy ops queue; one `context.sync()` sends the batch | No: the proxies are the only surface | Typed and huge: the `Excel` namespace is 69k lines and 193 classes, but Table + TableColumn + TableRow have 31 methods | First-party; docs updated 2026-06 |
| *ExcelJS (inspiration)* | TS typings over JS, `.xlsx` files | Workbook, Worksheet, **Table**, Row, Column, Cell | Yes, in the file | Nothing is remote; edits mutate an in-memory model. `Table` edits are cached until `table.commit()` | `.model` objects are exposed (for example `Worksheet.model`) | 2,034-line `index.d.ts` with 118 declarations; 24.5k-word README | 15.5k stars; v4.4.0; last push 2025-01-21 |

Baseline for comparison: the raw typed client `@googleapis/sheets@18.0.1` has a 6,785-line `v4.d.ts` with 262 `Schema$` interfaces, and `Schema$Request` alone has 69 optional request kinds, tables included.

## Evidence

### Apps Script `SpreadsheetApp`

- The class index lists no `Table` class. The only table-named classes are `DataSourceTable`/`DataSourceTableColumn` (Connected Sheets), `PivotTable` and `EmbeddedTableChartBuilder` ([reference index](https://developers.google.com/apps-script/reference/spreadsheet)). `Sheet` has no method with "Table" in its name ([Sheet](https://developers.google.com/apps-script/reference/spreadsheet/sheet)).
- The [Apps Script release notes](https://developers.google.com/apps-script/release-notes) have no Spreadsheet-service entry in 2025 or 2026. The last one is 2024-12-09 (`getSheetById()`). Sheets API tables arrived on 2025-04-29, and Apps Script users reach them through the Advanced Sheets service with raw `addTable` requests (for example [tanaikech's gist](https://gist.github.com/tanaikech/fe567e1f397dc4a8d88212f91e779dbd)).
- Batching: `SpreadsheetApp.flush()` "Applies all pending Spreadsheet changes. Spreadsheet operations are sometimes bundled together to improve performance, such as when doing multiple calls to Range.getValue()" ([SpreadsheetApp](https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app)). The bundling is internal. You can't see it, shape it, or turn it into one atomic `batchUpdate`.
- Raw beside the helpers: the [Advanced Sheets service](https://developers.google.com/apps-script/advanced/sheets) exposes `Sheets.Spreadsheets.batchUpdate` next to `SpreadsheetApp`.
- Surface: `@types/google-apps-script@2.0.13`, `google-apps-script.spreadsheet.d.ts`, has 2,607 lines, 68 interfaces and 1,404 method signatures. `Range` covers lines 1664-2053 with 245 signatures, `Sheet` 2054-2246 with 157, and `Spreadsheet` 2247-2399 with 133. The docs list 22 methods on `SpreadsheetApp`.

### gspread (`burnash/gspread` @ `7ca71ea`)

- Classes: `Client` (`client.py:23`), `Spreadsheet` (`spreadsheet.py:22`), `Worksheet` (`worksheet.py:161`), `Cell` (`cell.py:14`). There is no Row, Column or native Table class. `utils.TableDirection`/`find_table` (`utils.py:172`, `:1039`) expands from a cell until it reaches an empty one. It is not a Sheets table, and `tableId`/`addTable` appear nowhere.
- Each helper sends its own request. `worksheet.py` has 35 `.batch_update(` call sites and 12 `values_*` call sites, each called eagerly. Batching needs a dedicated method: `batch_update` (`worksheet.py:1264`), `batch_format` (`:1391`), `batch_clear` (`:2403`), `batch_merge` (`:2803`). There's no queue across method types.
- Raw: `Spreadsheet.batch_update(self, body: Mapping[str, Any])`, a "Lower-level method that directly calls" batchUpdate (`spreadsheet.py:92`).
- Surface: `def` counts are Worksheet 100, Spreadsheet 41, Client 16, Cell 5 and HTTPClient 21. The package is 7,416 lines. The hand-written docs are 4.5k words, and the API docs are autodoc from the source.

### google-spreadsheet (`theoephraim/node-google-spreadsheet` @ `d43ea3e`, npm 5.3.0)

- Classes in `dist/index.d.ts`: `GoogleSpreadsheetRow` (l.6), `GoogleSpreadsheetCellErrorValue` (584), `GoogleSpreadsheetCell` (596), `GoogleSpreadsheetWorksheet` (717) and `GoogleSpreadsheet` (1395). The only `tableId` is a filter-view field in the bundled types (`src/lib/types/sheets-types.ts:757`). There's no table helper.
- Cells batch: `saveUpdatedCells()` → `saveCells()` builds one `updateCells` per dirty cell and sends them all through `_makeBatchUpdateRequest` (`src/lib/GoogleSpreadsheetWorksheet.ts:271-292`).
- Rows don't. `GoogleSpreadsheetRow.save()` does a `values/<a1>` PUT (`GoogleSpreadsheetRow.ts:88-91`), and `delete()` sends a single `deleteRange` through `_makeSingleUpdateRequest` (`:108-116`). Worksheet ops such as `addRows`, `clearRows` and `delete` each send their own call (`GoogleSpreadsheetWorksheet.ts:413-453`, `:677`, `:1694`).
- Raw: `readonly sheetsApi: KyInstance`, documented as "can be used if unsupported sheets calls need to be made" (`GoogleSpreadsheet.ts:100-106`). `_makeBatchUpdateRequest(requests: any[])` is marked `@internal` but still ships in the `.d.ts` (l.1442).

### pygsheets (`nithinmurali/pygsheets` @ `7de8d48`, 2.0.5)

- Classes: `Client`, `Spreadsheet`, `Worksheet`, `DataRange`, `Cell`, `Address`/`GridRange` (`pygsheets/*.py`). The only `tableId` is inside the vendored discovery JSON.
- Batch mode: `Client.set_batch_mode` "will batch all custom requests and will combine them into single request … batch mode only caches sheetUpdate requests not value updates or clear requests" (`client.py:90-98`). `run_batch` returns inside its loop, so it sends only the first spreadsheet's queue and never clears the queue (`sheet.py:54-59`).
- Raw: `SheetAPIWrapper.batch_update(spreadsheet_id, requests)` (`sheet.py:61`).

### sheetfu (`socialpoint-labs/sheetfu` @ `d3a2ee5`)

- Classes: `Spreadsheet`, `Sheet`, `Range` (`model.py:22`, `:190`, `:310`) and `Table`/`Item` (`modules/table.py:7`, `:267`). `Table` wraps a Range whose first row is the header. `Item` gets and sets fields by header name, including notes, backgrounds and font colors.
- Batching is the closest to the target design. A range setter builds an `updateCells` request and either sends it or appends it to `batch_to.batches` (`model.py:440-473`). `Table.commit()` sends the whole queue as one `batchUpdate` (`modules/table.py:249-260`), and `Spreadsheet.commit()` does the same (`model.py:176-186`). Note that `commit` wraps the list twice as `{'requests': [self.batches]}`.
- Raw: `client.sheet_service` is the `googleapiclient` Sheets service (`client.py:41`).
- It's small (1,740 lines, 63 public methods), but untyped Python with no commits since 2022.

### ezsheets (`asweigart/ezsheets` @ `58d32e8`)

- Two classes, `Spreadsheet` and `Sheet` (`src/ezsheets/__init__.py:176`, `:535`), with row and column methods on `Sheet` (`getRow` at `:928` through `updateColumns` at `:1387`). Each write sends its own request. It has no passthrough and no tables.

### google-drive-ruby (`gimite/google-drive-ruby` @ `55b996b`)

- `Worksheet#[]=` (`lib/google_drive/worksheet.rb:184`) queues edits, and `save` sends them as one `spreadsheet.batch_update(@v4_requests)` (`:402-420`). `Spreadsheet#batch_update(requests)` is public (`spreadsheet.rb:126`). It has no tables, and its last commit was in 2021.

## Excel, for inspiration

### Office.js Excel JavaScript API (`@types/office-js@1.0.610`)

- Scopes: "A `Workbook` contains one or more `Worksheet` objects … A `Range` represents one cell or a block of contiguous cells", and "The Excel JavaScript API doesn't have a 'Cell' object or class" ([core concepts](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-core-concepts)). The typings declare `Workbook` (`index.d.ts:36716`, 22 methods), `Worksheet` (`:37245`, 27 methods), `Range` (`:38485`, 92 methods), `Table` (`:41696`, 14 methods and 19 properties), `TableColumnCollection` (`:41937`), `TableColumn` (10 methods), `TableRowCollection` (`:42176`) and `TableRow` (7 methods). The whole `declare namespace Excel` covers lines 26060-95516, with 193 classes.
- Table API: `sheet.tables.add("A1:D1", true)` creates a table, `table.rows.add(null, values)` appends rows, `table.columns.getItem("Merchant").getDataBodyRange()` addresses a column by name, and `getHeaderRowRange()`, `getDataBodyRange()` and `getTotalRowRange()` return the table's parts as ranges. "A `TableRow` object doesn't have an `id` property, so use the row position", while a `TableColumn`'s `id` "contains a unique key that identifies the column" ([tables](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-tables)).
- Batching: "Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands. When you call the `sync()` method … the queued commands are dispatched … These APIs are fundamentally batch-centric." Reads need an explicit `load('address')` before the next `sync()`. `getItemOrNullObject()` checks for existence without throwing, and `set({...})` sets nested properties from one object literal ([application-specific API model](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model)).
- Raw: there's no public request layer under the proxies. The typed object model is the only surface.

### ExcelJS (`exceljs/exceljs` @ `5bed18b`, npm 4.4.0)

- Scopes: `Workbook` (`index.d.ts:1707`), `Worksheet` (`:1094`), `Row` (`:523`), `Column` (`:603`), `Cell` (`:424`) and `Table` (`:1878`). The Worksheet has `addTable`, `getTable`, `removeTable` and `getTables` (`:1361-1373`).
- Table edits (`addRow`, `removeRows`, `addColumn`, …) snapshot the old footprint through `cacheState()`. `commit()` then validates the table, blanks the cells it no longer covers and re-stores it (`lib/doc/table.js:285-335`). The README calls `table.commit()` after every modification (README "Modifying Tables").
- It works only on files. Nothing is sent over a network, so batching doesn't apply.

### What a Sheets harness could borrow

1. **One queue per context, flushed once.** Office.js's proxies-then-`sync()` maps directly onto Sheets: scoped objects append `Schema$Request`s to a context, and `send()` becomes one `batchUpdate`. Unlike Office's `sync()`, a Sheets `batchUpdate` is atomic, and an unsent queue is exactly what a harness preview would inspect.
2. **A small table scope made of named range parts.** Office.js's header, data-body and total-row ranges, `columns.getItem(name)` and `rows.add(null, values)` fit in 31 methods. These map onto Sheets' `tableId`, the table's column properties, and `appendCells` with a `tableId`.
3. **Explicit `load` of named fields** corresponds to the Sheets API `fields` mask, which keeps reads cheap.
4. **`*OrNullObject` lookups** let an agent check that a sheet or table exists without try/catch.
5. **ExcelJS's footprint reconciliation on `commit()`.** When a table shrinks or moves, it clears the cells the table leaves behind, which is what a blank-row-safe table delete needs.
6. **Stay small.** Office.js's full surface (69k lines) is the cost to avoid. Its table subset is the size to aim for.

## What this means for the harness

1. **Table scope is unclaimed.** No class-based library models native tables, and Apps Script has shown no sign of adding one (no Spreadsheet-service release since 2024-12).
2. **The "queue then one batchUpdate" pattern has precedent** in sheetfu (`batch_to`/`commit`), google-drive-ruby (`save`) and pygsheets (`set_batch_mode`). All three are untyped and dormant or buggy, and none of them queues value writes together with structural requests.
3. **A typed raw request beside typed helpers is unclaimed.** Every escape hatch takes `any`, a `dict`, or a raw HTTP client. Typing it as `sheets_v4.Schema$Request` from `@googleapis/sheets` is available and unused.
4. **Parsimony is the agent-cost lever.** Apps Script's 1,404 signatures are the typed incumbent's cost. google-spreadsheet shows a typed scoped model can fit in a ~1.6k-line `.d.ts`, and sheetfu shows table-ish helpers can fit in 23 methods.

Verdict: open ground for Sheets. The nearest Sheets neighbours are google-spreadsheet (typed, small, maintained, but no tables and ad hoc batching) and sheetfu (a table abstraction and explicit batch commit, but untyped, not native tables, and dead). The closest overall design is Office.js's Excel API, which shows what the target looks like once it is built. A Sheets harness can copy its queue-and-sync model and its table subset while leaving the rest of its breadth behind.
