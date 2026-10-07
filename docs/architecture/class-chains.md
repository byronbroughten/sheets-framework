# The sheet and column class chains

Map fragment. Sibling headings live in this folder.

The shape every tier's sheet and column classes share, the extra Table-level class the Raw tier needs, and the guard that keeps the widened column type from collapsing to `never`.

## The three-level shape

Every tier's sheet and column classes have the same three-level shape: a base class, an abstract common class, and the two concrete Meta/primary classes under it.

```
ColumnBaseNamed<SN, CN>          // the base: columnName, schema, columnNamedProps
  └─ ColumnCommonNamed<SN, CN>   // abstract; adds columnId
       ├─ ColumnNamed<SN, CN>
       └─ ColumnMetaNamed<SN, CN>
```

There is no base class for the primary column alone: a class that wants to sit beside `ColumnNamed` rather than under it extends `ColumnBaseNamed` and reaches the column through a getter — which is what `GenericTableOperator` does one level up (`extends TableBaseNamed`), and what the column-scoped endpoint classes in `06_API` already do. Spreadsheet-scoped Operators that own config-sync state extend `SpreadsheetBaseOperator` instead (`ConfigCoordinator`), which is the Named spreadsheet base plus `configSyncState`.

## Row and cell base classes

**Rows and cells have base classes too, and an operator may hang off either.** The Named tier's `ClassBases/` holds `SpreadsheetBaseNamed`, `TableBaseNamed` (adds `sheetName`), `RowBaseNamed` (adds `rowIndex`) and `CellBaseNamed` beside the column chain's classes. The chains above show sheets and columns because those two needed explaining, not because they are the set an operator may extend. The house style's class-shape reasoning covers choosing between them.

## The Raw tier's two Table-level classes

The Raw tier needs **two** Table-level classes above its concrete pair, and both earn their place:

```
TableBaseRaw                     // the address (a tableId, or a sheet GID until Meta is retired), sheet and Table state, schema
  ├─ TableCommonRaw              // abstract; ss, the Table's identity and bounds, fetch ranges, the Table write queue
  │    ├─ TableRaw
  │    └─ SheetMetaRaw
  ├─ RowBaseRaw
  └─ ColumnBaseRaw
```

`TableCommonRaw` holds the Table's identity and geometry: `tableId`, `name`, the bounds, `dataRowCount`, `columnCount` and the stale flag. Its state is `TableStateRaw`, keyed by `tableId`, which holds the Table's properties, rows, columns and fetch and write queues; `SheetStateRaw` keeps only per-sheet facts. Identified's chain is `TableBaseIdentified` → `TableCommonIdentified` → `TableIdentified`. Meta and the Named classes keep their Sheet names.

**Raw addresses a Table by its live `tableId`** (`ss.table(tableId)`). Until Meta is retired, `ss.sheetMeta(gid).primary` resolves the sheet's one Table. Before that Table's properties are fetched, it queues into the sheet's `tableBeforeProperties` state, and the Table adopts it when its properties arrive.

`TableBaseRaw` cannot hold `TableCommonRaw`'s members, because the row and column base classes hang off it: `RowCommonRaw` already declares a `writes` of an incompatible type and `CellRaw` a `gatherFetchRange` of a different signature, so either would be an illegal override, and the rest would be inherited by classes with no use for them. Don't fold `TableCommonRaw` back into `TableBaseRaw`. `ss` is declared `abstract` on `TableCommonRaw` and implemented on each concrete class — importing `SpreadsheetRaw` as a value there would close a module-init cycle through the two subclasses' `extends` clauses, which is the crash class described under [The schema classes](./schema-classes.md). `RowCommonRaw → TableRaw → RowRaw extends RowCommonRaw` has that shape latent today: it loads only because every current entry point reaches `TableRaw` or `RowRaw` before `RowCommonRaw`. Rows, columns, and cells reach their Table through a `table` getter; there is no `sheet` getter on those classes. Identified extends Raw's spreadsheet base, so one instance holds both tiers' spreadsheet-level state; the members carry the tier suffix (`spreadsheetStateRaw`, `spreadsheetStateIdentified`) rather than overriding one name.

## The widened instantiation never collapses to `never`

**The widened instantiation must never collapse to `never` — and today it doesn't.** The hazard is real but guarded: indexing a union of sheets by a union of column names naively requires the column name to key *every* sheet, which would yield `never`, so `ColumnValueName` (in `columnConfigsTypes.ts`) is deliberately written to distribute over the sheet name instead. That is why it looks the way it does; it is not a description of a current defect. Type assertions in `SpreadsheetSchema.test.ts` pin both ends — exact literal on the named path, full union and specifically not `never` on the widened one — and fail `npm run tsc` if either degrades. `tsc` passing is not by itself evidence that precision survived, which is why those assertions are not optional. Neither is an assignment — the house style's rule is "Verify a type-level claim with an identity check."
