# The sheet and column class chains

Map fragment. Sibling headings live in this folder.

The shape every tier's sheet and column classes share, the extra sheet-level class the Raw tier needs, and the guard that keeps the widened column type from collapsing to `never`.

## The three-level shape

Every tier's sheet and column classes have the same three-level shape: a base class, an abstract common class, and the two concrete Meta/primary classes under it.

```
ColumnBaseNamed<SN, CN>          // the base: columnName, schema, columnNamedProps
  └─ ColumnCommonNamed<SN, CN>   // abstract; adds columnId
       ├─ ColumnNamed<SN, CN>
       └─ ColumnMetaNamed<SN, CN>
```

There is no base class for the primary column alone: a class that wants to sit beside `ColumnNamed` rather than under it extends `ColumnBaseNamed` and reaches the column through a getter — which is what `GenericSheetOperator` does one level up (`extends SheetBaseNamed`), and what the column-scoped endpoint classes in `06_API` already do. Spreadsheet-scoped Operators that own config-sync state extend `SpreadsheetBaseOperator` instead (`ConfigCoordinator`), which is the Named spreadsheet base plus `configSyncState`.

## Row and cell base classes

**Rows and cells have base classes too, and an operator may hang off either.** The Named tier's `ClassBases/` holds `SpreadsheetBaseNamed`, `SheetBaseNamed` (adds `sheetName`), `RowBaseNamed` (adds `rowIndex`) and `CellBaseNamed` beside the column chain's classes. The chains above show sheets and columns because those two needed explaining, not because they are the set an operator may extend. The house style's class-shape reasoning covers choosing between them.

## The Raw tier's two sheet-level classes

The Raw tier needs **two** sheet-level classes above its concrete pair, and both earn their place:

```
SheetBaseRaw                     // sheetGid, sheet state, schema
  ├─ SheetCommonRaw              // abstract; ss, activeTable, fetch-range gatherers, sheet change queue, fullTableColIndexes
  │    ├─ SheetRaw
  │    └─ SheetMetaRaw
  ├─ RowBaseRaw
  └─ ColumnBaseRaw

ActiveTableRaw                   // collaborator beside the chain, not a SheetBaseRaw subclass
```

`SheetBaseRaw` cannot hold `SheetCommonRaw`'s members, because the row and column base classes hang off it: `RowCommonRaw` already declares a `changesToSave` of an incompatible type and `CellRaw` a `gatherFetchRange` of a different signature, so either would be an illegal override, and the rest would be inherited by classes with no use for them. Don't fold `SheetCommonRaw` back into `SheetBaseRaw`. `ss` is declared `abstract` on `SheetCommonRaw` and implemented on each concrete class — importing `SpreadsheetRaw` as a value there would close a module-init cycle through the two subclasses' `extends` clauses, which is the crash class described under [The schema classes](./schema-classes.md). `RowCommonRaw → SheetRaw → RowRaw extends RowCommonRaw` has that shape latent today: it loads only because every current entry point reaches `SheetRaw` or `RowRaw` before `RowCommonRaw`. Rows, columns, and cells reach the Table through `this.sheet.activeTable`; there is no Table getter on those classes and no `sheet` getter on `RowBaseRaw`. Identified extends Raw's spreadsheet base, so one instance holds both tiers' spreadsheet-level state; the members carry the tier suffix (`spreadsheetStateRaw`, `spreadsheetStateIdentified`) rather than overriding one name.

## The widened instantiation never collapses to `never`

**The widened instantiation must never collapse to `never` — and today it doesn't.** The hazard is real but guarded: indexing a union of sheets by a union of column names naively requires the column name to key *every* sheet, which would yield `never`, so `ColumnValueName` (in `columnConfigsTypes.ts`) is deliberately written to distribute over the sheet name instead. That is why it looks the way it does; it is not a description of a current defect. Type assertions in `SpreadsheetSchema.test.ts` pin both ends — exact literal on the named path, full union and specifically not `never` on the widened one — and fail `npm run tsc` if either degrades. `tsc` passing is not by itself evidence that precision survived, which is why those assertions are not optional. Neither is an assignment — the house style's rule is "Verify a type-level claim with an identity check."
