# Sheets framework

A Google Sheets spreadsheet that a person operates directly, with an Apps Script layer built on this framework reacting to their edits. The vocabulary below is the language of that operator-facing surface — what a person clicks, and what the sheet tells them back — in every app built on the framework. An app's own glossary refines these terms and adds its domain's.

## Language

### Sheet layout

**Table**:
The Google Table (Insert > Table) laid over a sheet's data. Every sheet with **Let api access** must have exactly one, starting on the Table header row in the first column, with at least one data row (a **blank row** counts). What a run does about a moved or extra Table: [`docs/architecture/table-placement.md`](./docs/architecture/table-placement.md).
_Avoid_: range, data range, grid

**Table reference**:
A formula that names a Table column by the Table's name and the column header — `test[Number]`, usually wrapped in `SINGLE(...)` when one cell is wanted — so the formula stays readable when columns move. It is not an A1 address like `$C5`.
_Avoid_: structured reference, A1, cell address

**Column ID row**:
The bookkeeping row of generated column identifiers, above the other two bookkeeping rows. You never edit it by hand; the app fills a blank when a Table column has none.
_Avoid_: ID row, metadata row, row 1

**Column-group heading**:
The bookkeeping row of group names, between the column ID row and the action row. You never edit it by hand.
_Avoid_: group row, section header

**Action row**:
The row above the Table header row where an endpoint is triggered. Most of its cells are empty, and a cell may hold text used as a label. Only the cells wired to an endpoint hold a checkbox, and ticking one of those is what asks the spreadsheet to do something, one endpoint per column.
_Avoid_: control row, button row, trigger row

**Table header row**:
The row of column titles you read across the top of a sheet's data, directly above the first data row, and the row the Table starts on. Three bookkeeping rows sit above it that you never edit by hand: the column ID row, the column-group heading, and the action row.
_Avoid_: header row, title row, top row, row 1

**First data row**:
The first row of the Table's data, always the row immediately below the Table header row. That index is not stored separately.
_Avoid_: data start, top data row, row 5

**Layout value**:
One of Spreadsheet Config's cells saying where the bookkeeping rows and the Table sit and how IDs are written. All but ID header and Name header are fixed for now.
_Avoid_: layout setting, layout constant, spreadsheet config value

**Blank row**:
A data row with nothing in any of the columns you fill in yourself, which the app leaves when it deletes everything on a sheet. The next row the app adds goes into it. Why, and how: [`docs/architecture/blank-row.md`](./docs/architecture/blank-row.md).
_Avoid_: empty row, placeholder row, spare row

**ID prefix**:
The short, readable code every row ID and column ID on a sheet begins with, so you can tell at a glance which sheet an ID belongs to. The app gives one to every sheet it knows about, taken from the tab title. No two sheets share one, and renaming the tab doesn't change it.
_Avoid_: sheet prefix, ID code

**Let api access**:
The Sheet Config checkbox that says this tab is one the app knows about — not every tab, and not every catalogue row on Sheet Config.
_Avoid_: enabled sheet, API sheet, known sheet

**Edit protection**:
Any protection the app finds on a sheet: an edit warning, an edit lock, or one it can't read as either, which it leaves alone.
_Avoid_: protected range (that is Google's name for the API object)

**Edit warning**:
A prompt Sheets shows anyone, the owner included, before they change a cell the app depends on or the shape of a tab it depends on (renaming it, deleting it, inserting a column); the change still goes through if they confirm.
_Avoid_: warning (that is a run state), protection

**Editable range**:
A part of a warned floor tab left free of its edit warning, where a person's edit sticks. It skips a self-describing row's declared cell, so an edit there prompts.
_Avoid_: hole, unprotected range (that is Google's API field), exception

**Edit lock**:
A cell only the editors it names can change; a lock that names none stops nobody.
_Avoid_: protection, lock

**Config-sheet floor**:
Everything on the four config sheets that the app guarantees and restores on each config sync: tab titles, Table names, headers, column IDs, column-group headings, column types, Spreadsheet Config's Table menu space data cell, and the framework endpoints' columns. Edits to it are overwritten and reported ([`docs/generated-data/config-sheet-floor.md`](./docs/generated-data/config-sheet-floor.md)).
_Avoid_: minimum floor, minimum headers, floor sheet

**Floor seed**:
The app's own declaration of what the config-sheet floor looks like: its structure, plus Table menu space's data value. A floor tab or column keeps the identity it was created with. A tab it creates also gets its **seeded values**.
_Avoid_: template, default config

**Self-describing row**:
A Sheet Config or Column Config row describing a config-sheet floor tab or floor column; its declared cell is enforced from the floor seed.
_Avoid_: config-about-config row

**Seeded value**:
A cell or column the floor seed fills once, when it creates the tab or Table, and never restores — such as Spreadsheet Config's **layout values** or Value Config's example column.
_Avoid_: initial value (too close to Custom default value)

**Floor notice**:
A message that pops up when someone changes the config-sheet floor in a way the next config sync will reverse, saying what will happen and what they can do about it.
_Avoid_: toast, warning (that is a run state), alert

### Endpoints

**Endpoint**:
A unit of work the spreadsheet can be asked to do, wired to one column and triggered by a checkbox in that column's action row. Any column can be the one; it declares for itself which other columns the framework should manage on its behalf.
_Avoid_: handler, command, action

**Framework endpoint**:
An endpoint the framework itself provides, wired to a column on Spreadsheet Config. The config-sheet floor restores its column on every config sync, so it is present in every spreadsheet whatever endpoints the app adds, and an app's own endpoint can't be wired to a config-sheet column, so the two never share a column.
_Avoid_: base endpoint, built-in endpoint, core endpoint, system endpoint

**Runner**:
An endpoint whose entry checkbox is a run button: ticking it starts the work, and the box clears itself immediately. An endpoint is one unless it says it also runs on unticking.
_Avoid_: job, task, trigger

**Two-way endpoint**:
An endpoint whose entry checkbox is the input rather than a button — it runs on ticking and on unticking, is told which way it went, and the box stays where the operator left it.
_Avoid_: toggle, switch

**Selector**:
A column of checkboxes an endpoint may declare, naming the rows one run is about; the run acts on and reports into those rows only. A successful run **consumes** its selection unless the endpoint **retains its selection**, and an endpoint may declare that it **requires one row**. What the operator sees: [`docs/architecture/endpoint-dispatch.md`](./docs/architecture/endpoint-dispatch.md#what-the-operator-sees-of-a-selector).
_Avoid_: toggle, filter

**Feedback column**:
A column an endpoint declares for the framework to write into on its behalf, rather than one the endpoint's own work fills: the start-time column and the run-status column. Each is optional; an endpoint that declares neither reports nothing.
_Avoid_: output column, status column

**Run status**:
The sentence an endpoint writes for the operator to read: that it is running, that it succeeded — in its own words if it has any — or what went wrong. It is written into the run-status cell of every row the run is about, and a run may give a particular row a sentence of its own instead.
_Avoid_: error message, log, result

**Run state**:
Which of an endpoint's four conditions its last run is in — running, success, warning or failure. A run state is always a run status message and a colour together, so the two can never disagree, and both feedback columns are painted in it, so the state is visible whichever of them an endpoint declares. Every row the run is about carries one, and a run may put a different one on a particular row.
_Avoid_: run outcome, status code

**Running**:
Work has begun and has not reported back. A run killed mid-flight stays here, which is how "died" is distinguishable from success, warning and failure alike.
_Avoid_: in progress, pending, processing

**Warning**:
The run committed its work, and something about it wants your attention — most often that some of the rows it was about went through and some did not. It is orange and always carries a sentence of its own.
_Avoid_: partial, incomplete, soft failure

**Run report**:
What an endpoint hands back when its work is done: nothing, a sentence, a run state with a sentence, or a set of rows that differ from the rest, each with the state and sentence it gets. A run that **fails by throwing** is not a run report: its work is abandoned and every row goes red.
_Avoid_: result, outcome, return value

**Start time**:
When a run began, written once into every row the run is about and never rewritten, so elapsed time stays readable while a slow run is still going.
_Avoid_: finished at, completion time, duration

### Columns

**Column type**:
What a column holds, as the operator declares it in the column's own type menu in Sheets — currency, date, checkbox, text, and the rest. The app trusts it over anything it could work out for itself from the data, including the column's number format.
_Avoid_: data type, format, value type

**Number format**:
What Format > Number says on a column's first data row — currency, date, number, plain text, and the rest. When the type menu is silent, a format the app knows counts as a declaration only if that row's value is **compatible** with it ([`docs/generated-data/column-configs.md`](./docs/generated-data/column-configs.md#how-valuename-is-resolved)).
_Avoid_: column type, value type, cell type, permissible

**Checkbox column**:
A column the operator made a checkbox: the type menu says Checkbox, or Insert > Checkbox put BOOLEAN data validation on the Table column or the first data row. A row nobody has touched counts as unchecked rather than as blank, and only such a column can be an endpoint's selector.
_Avoid_: boolean column, tickbox column, flag column

**Empty value allowed**:
A box you tick against a column in Column Config to say that a blank in it is a real answer rather than something missing. Unticked, the app stops and names the cell whenever it reads a blank there; ticked, it hands the blank on to whatever asked for it, and that work has to say what a missing value means.
_Avoid_: nullable, optional column, blank allowed

**Untyped**:
Said of a column whose type menu, checkbox validation, and first-data-row number format all tell the app nothing about what it holds — left on Automatic with no format the app maps and no Insert > Checkbox, or a dropdown that no Value Config rule backs. The app guesses its type from the column's top value.
_Avoid_: unset, automatic, missing type

**Name column**:
A column of names people type to refer to a row, such as a customer's or a project's name, on any sheet whose Table header row holds Spreadsheet Config's Name header. An endpoint finds a row by what's in it; a name that matches no row, or several, is the endpoint's to report, and the app enforces nothing else about it.
_Avoid_: key column, label column, title column

**Serial date**:
A date as Sheets stores it: the count of days since 30 December 1899. Adding or comparing serial dates doesn't depend on a timezone; knowing which one is today does, and the spreadsheet's own timezone decides that.
_Avoid_: date number, day serial, JS date

