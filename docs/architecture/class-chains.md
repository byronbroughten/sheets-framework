# The Table and column class chains

Map fragment. Sibling headings live in this folder.

The shape every tier's Table and column classes share, the extra Table-level class the Raw tier needs, and the guard that keeps the widened column type from collapsing to `never`.

## The three-level shape

Every tier's Table and column classes have the same three-level shape: a base class, an abstract common class, and the concrete class under it.

```
ColumnBaseNamed<TN, CN>          // the base: columnName, schema, columnNamedProps
  └─ ColumnCommonNamed<TN, CN>   // abstract; adds columnId
       └─ ColumnNamed<TN, CN>
```

There is no base class for the concrete column alone: a class that wants to sit beside `ColumnNamed` rather than under it extends `ColumnBaseNamed` and reaches the column through a getter — which is what `GenericTableOperator` does one level up (`extends TableBaseNamed`), and what the column-scoped endpoint classes in `06_API` already do. Spreadsheet-scoped Operators that own config-sync state extend `SpreadsheetBaseOperator` instead (`ConfigCoordinator`), which is the Named spreadsheet base plus `configSyncState`.

## Row and cell base classes

**Rows and cells have base classes too, and an operator may hang off either.** The Named tier's `ClassBases/` holds `SpreadsheetBaseNamed`, `TableBaseNamed` (adds `tableName`), `RowBaseNamed` (adds `rowIndex`) and `CellBaseNamed` beside the column chain's classes. The chains above show Tables and columns because those two needed explaining, not because they are the set an operator may extend. The house style's class-shape reasoning covers choosing between them.

## The Raw tier's two Table-level classes

The Raw tier needs **two** Table-level classes above its concrete pair, and both earn their place:

```
TableBaseRaw                     // the address (a tableId, or the GID ss.tableOnSheet resolves), sheet and Table state, schema
  ├─ TableCommonRaw              // abstract; ss, the Table's identity and bounds, fetch ranges, the Table write queue
  │    ├─ TableRaw
  │    └─ TableColumnResolverRaw
  ├─ TableProfileRaw             // reads only, so no queued writes
  ├─ RowBaseRaw
  └─ ColumnBaseRaw               // ColumnRaw and ColumnProfileRaw hang off it
```

`TableCommonRaw` holds the Table's identity and geometry: `tableId`, `name`, the bounds, `dataRowCount`, `columnCount` and the stale flag. Its state is `TableStateRaw`, keyed by `tableId`, which holds the Table's properties, rows, columns and fetch and write queues; `SheetStateRaw` keeps only per-sheet facts. Identified's chain is `TableBaseIdentified` → `TableCommonIdentified` → `TableIdentified`, and Named's is `TableBaseNamed` → `TableCommonNamed` → `TableNamed`.

**Raw addresses a Table by its live `tableId`** (`ss.table(tableId)`). Raw's `ss.tableOnSheet(gid)` resolves a sheet's Table by GID: its only one, else the one its configs record; before that Table's properties are fetched, it queues into the sheet's `tableBeforeProperties` state. Identified reaches every managed Table by its recorded `tableId` (Named passes it as a `TableAddressIdentified`); Identified's `ss.tableOnSheet(gid)` means the one Table the configs record on that sheet and throws on a sheet they record several on. Before a recorded Table's properties arrive, Raw queues it into the sheet's `tablesBeforePropertiesById`.

`TableBaseRaw` cannot hold `TableCommonRaw`'s members, because the row and column base classes hang off it: `RowCommonRaw` already declares a `writes` of an incompatible type and `CellRaw` a `gatherFetchRange` of a different signature, so either would be an illegal override, and the rest would be inherited by classes with no use for them. Don't fold `TableCommonRaw` back into `TableBaseRaw`. `ss` is declared `abstract` on `TableCommonRaw` and implemented on each concrete class — importing `SpreadsheetRaw` as a value there would close a module-init cycle through the two subclasses' `extends` clauses, which is the crash class described under [The schema classes](./schema-classes.md). `RowCommonRaw → TableRaw → RowRaw extends RowCommonRaw` has that shape latent today: it loads only because every current entry point reaches `TableRaw` or `RowRaw` before `RowCommonRaw`. Rows, columns, and cells reach their Table through a `table` getter; there is no `sheet` getter on those classes. Identified extends Raw's spreadsheet base, so one instance holds both tiers' spreadsheet-level state; the members carry the tier suffix (`spreadsheetStateRaw`, `spreadsheetStateIdentified`) rather than overriding one name.

## The Sheet container

**The Sheet container is `SheetRaw extends SpreadsheetBaseRaw` and `SheetNamed extends SpreadsheetBaseNamed`, with no base, common or Identified class**, since nothing extends them; bring the bases back if sheet-scoped classes multiply. `ss.sheet(gid)` reaches it at Raw, `ss.sheet(title)` at Named, and `table.sheet` at both; `TableIdentified` has no `sheet` and goes through `raw`. It holds the tab's own facts: its `title`, its grid's `rowCount` and `columnCount` and their `appendDimension` totals, the whole-sheet protections, and the conditional-format and protection lists, the latter two in `SheetRaw/` collaborators; the Table keeps only the actions bounded to its own data range. It also lists its managed Tables: `sheet.tableIds` at Raw and `sheet.tableNames` at Named read `tableConfigs`, so a Table the configs don't record isn't listed.

## The widened instantiation never collapses to `never`

**The widened instantiation must never collapse to `never` — and today it doesn't.** The hazard is real but guarded: indexing a union of Tables by a union of column names naively requires the column name to key *every* Table, which would yield `never`, so `ColumnValueName` (in `columnConfigsTypes.ts`) is deliberately written to distribute over the Table name instead. That is why it looks the way it does; it is not a description of a current defect. Type assertions in `SpreadsheetSchema.test.ts` pin both ends — exact literal on the named path, full union and specifically not `never` on the widened one — and fail `npm run tsc` if either degrades. `tsc` passing is not by itself evidence that precision survived, which is why those assertions are not optional. Neither is an assignment — the house style's rule is "Verify a type-level claim with an identity check."
