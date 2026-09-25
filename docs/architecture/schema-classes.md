# The schema classes

Map fragment. Sibling headings live in this folder.

The four schema classes: where they live, why the split is safe from module-init cycles, what each answers, and the two ways to build one.

## Where they live and why the split is safe

The Schema tier root, `src/01_SpreadsheetSchema/`, holds the four schema classes, one per file, below Raw so every tier can use them. Beside them, the `uniformRows` bundle reads the uniform-row indexes from the spreadsheet config. `SpreadsheetBaseSchema` wraps it, and the uniform-row constructors import it directly because they need an index before `super()`, when `this.schema` isn't reachable yet. All three concrete classes extend `SpreadsheetBaseSchema`, which imports none of them. That is what makes the split safe. `extends` is evaluated at module init, so a base class that reaches its own subclass through a value import crashes with "Class extends value undefined" whenever the base is the first module of its cycle to load. Neither `tsc` nor the tests catch that. `SpreadsheetSchema.sheetByName` builds a `SheetSchema`, so it lives on `SpreadsheetSchema` rather than on the shared base.

## Sibling classes may import each other

The converse is worth knowing, because it looks like the same hazard and isn't: **two sibling classes may import each other as values and instantiate each other in getter bodies.** `SheetSchema`/`ColumnSchema`, `SheetRaw`/`SheetMetaRaw` and `ColumnRaw`/`ColumnMetaRaw` already do, in both directions. Only `extends` runs at module init, so a cycle whose imports are used exclusively inside method and getter bodies is safe. When a cycle would close through an `extends` clause, move the shared members to a base that imports no subclass, as `SpreadsheetBaseSchema` does, or use an `abstract` member implemented on each subclass, as `SheetCommonRaw` does with `ss`.

## What each class answers

| Class                   | Answers                                                 | Reached from                        |
| ----------------------- | ------------------------------------------------------- | ----------------------------------- |
| `SpreadsheetBaseSchema` | Uniform-row indexes, ID encode/decode, layout constants | Every Raw-tier class                |
| `SpreadsheetSchema`     | The sheet list and navigation to a sheet's schema       | Spreadsheet-level classes           |
| `SheetSchema<SN>`       | A sheet's traits, its column IDs and names              | Sheet-level and row-level classes   |
| `ColumnSchema<SN,CN>`   | A column's value name, validation, default, full name   | Column-level and cell-level classes |

`SheetSchema` and `ColumnSchema` are **siblings**, both extending `SpreadsheetBaseSchema`, as does `SpreadsheetSchema`. `ColumnSchema` does *not* extend `SheetSchema` — it reaches its sheet through a `sheet` accessor. That's forced: one accessor named `schema` declared at the Raw root means every narrowing override must be assignable to what the root declares, and the consuming class tree branches (the sheet class, the row base and the column base are siblings under a shared sheet-scoped base), so the two need only be assignable to `SpreadsheetBaseSchema`, never to each other. The Raw root declares `schema` as `SpreadsheetBaseSchema`, and each concrete spreadsheet class (`SpreadsheetRaw`, `SpreadsheetIdentified`, `SpreadsheetNamed`, `Api`) narrows it to `SpreadsheetSchema`.

## One `schema` accessor

**One accessor.** Every class that has a schema exposes it as `schema`, narrowed to its level. There is no `baseSchema`, `ssSchema`, `columnSchema` or `sheetSchema` — if you find one, it's a leftover.

## Two addressing modes

**Two addressing modes, one class.** Both are entry points rather than separate hierarchies, and each resolves all of its coordinates eagerly at construction so later lookups use whichever index is cheapest:

| Built from | Entry point                                | Types                                               |
| ---------- | ------------------------------------------ | --------------------------------------------------- |
| Sheet name | `SheetSchema.fromSheetName(sheetName)`     | Full literal precision — column names autocomplete  |
| Sheet GID  | `SheetSchema.fromSheetGid(sheetGid)`       | Widened defaults, as Identified-tier callers expect |
| Both names | `ColumnSchema.fromColumnName(sn, cn)`      | Value name resolves to its exact literal            |
| GID + ID   | `ColumnSchema.fromColumnId(gid, columnId)` | Value name is the full union                        |

The statics are named distinctly per class because static members are inherited, so same-named helpers of different arities would collide. Navigation avoids the bare name `sheet` for the same reason it's reserved on `ColumnSchema`: `SpreadsheetSchema.sheetByName`/`.sheetByGid`, `SheetSchema.columnByName`/`.columnById`.
