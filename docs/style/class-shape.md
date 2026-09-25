# Class shape

Style fragment. The one-line rules live in [`docs/style.md`](../style.md); this file holds the reasoning and the worked examples. The general class-shape reasoning (member order, helpers, dead scaffolding) is in `@byronbroughten/config`'s `docs/style/class-shape.md`.

One heading per class-shape rule, most of them with the refactor that produced it: grep `^## ` for the rule you're applying and read that section.

## Coordinator classes

When code coordinates other stateful objects (other Operators, a Spreadsheet), write it as a **coordinator class** extending the tier's Base class (matching `SheetConfigOperator`).

This was mined from the framework tiers. It does **not** govern business endpoints: an endpoint is a plain entry with module-private helpers, matching the endpoints already in the registry. Reach for a class only once a file is unwieldy or its logic finds a second caller.

**What unwieldy meant, the one time it was crossed.** `buildLedger.ts` was 303 lines holding eighteen free functions, and almost every one of them existed to thread a collaborator through: the spreadsheet appeared in six of their signatures, the occupancy id in three, and a map of charges keyed by id in one more (#22). The number is the threshold's first data point, not its definition — what made it unwieldy was the threading, and the win was signatures rather than lines, since the body came out about the same length as `OccupancyLedgerOperator`.

- **`init`** is how a caller outside the class builds one; **`new`** is how a class builds its own collaborators from props already on `this`. `init` takes whatever the caller already holds — nothing at all for an entry point that starts a run (`ConfigCoordinator.init()`), or a live collaborator whose state must be shared, as when an endpoint action builds an operator from the `ss` it was handed — and assembles the props itself. A collaborator reached from props already on `this` skips `init` and is constructed directly in a getter: `new ColumnConfigOperator(this.operatorProps)` when the collaborator needs config-sync state, otherwise `new SpreadsheetNamed(this.spreadsheetNamedProps)`. Either way the real constructor just takes a `props` object.
- Collaborators (`ss`, `sheetConfigOperator`, `schema`, etc.) are lazy getters built from shared props on `this`.

```ts
export class ConfigCoordinator extends SpreadsheetBaseOperator {
  constructor(props: SpreadsheetNamedProps) {
    super({
      ...props,
      configSyncState: SpreadsheetBaseOperator.initConfigSyncState(),
    });
  }
  static init(): ConfigCoordinator {
    return new ConfigCoordinator(
      SpreadsheetBaseNamed.initSpreadsheetNamedProps(),
    );
  }
  get sheetConfigOperator() {
    return new SheetConfigOperator(this.operatorProps);
  }
  get columnConfigOperator() {
    return new ColumnConfigOperator(this.operatorProps);
  }
  syncAndFlushConfigSheets() {
    this.sheetConfigOperator.fetchAndUpdateAll();
    this.columnConfigOperator.fetchAndUpdateColumnConfig();
    this.ss.batchUpdateGSheets();
  }
}
```

Callers reach `orchestrator.sheetConfigOperator` on the instance. The method returns nothing.

## Split a coordinator when its helpers share nothing

A coordinator splits when its private helpers form groups that share no helpers with each other; each group becomes a collaborator. Method count doesn't decide it, because a long class whose helpers all call each other has no seam to cut along, and a short one with two unrelated jobs does. `SpreadsheetRaw` had 26 private helpers in three such groups: finishing a gathered fetch, checking that each sheet's Table is where the layout requires, and turning queued changes into a batch update and sending it. A reader looking for one job had to scroll through the other two (#67).

- **Collaborators follow the coordinator rules above.** They extend the tier's Base class, are built with `new` from the coordinator's props, and are reached through lazy getters (`fetcher`, `flusher`). A collaborator that needs another reaches it the same way: `SpreadsheetFetcherRaw` builds its own `tableValidator`, because finishing a fetch is what decides which Tables to judge.
- **The split is internal.** The coordinator keeps each public method as a one-line delegation, so no caller and no test changes, and the collaborators get no tests of their own: the coordinator's tests already cover them.
- **Collaborators live in a subfolder named after their coordinator** (`02_SpreadsheetRaw/SpreadsheetRaw/`), so the tier root lists only the classes callers use. A base class serves a whole chain of classes, often across tiers, so it stays in `ClassBases/` and never goes in one coordinator's subfolder.
- **A collaborator reaches back to the coordinator's shared surface through a getter named for it**: `ss` for a spreadsheet coordinator, as `SpreadsheetFlusherRaw` does, and `sheet` for a sheet one, as `SheetEditProtectionsRaw` does, rather than copying `sheet(sheetGid)` into each collaborator. The import cycle this makes is safe because the coordinator is only used inside method bodies, never in an `extends` clause.

## An Operator extends a `*BaseNamed` and reaches its subject through a getter

Whatever an Operator operates on — a sheet, a column — it extends that thing's `*BaseNamed` class and adds methods suited to that data structure. It does **not** extend the concrete class it works through, and it doesn't take one as a constructor argument: the subject is a lazy collaborator getter built from the props already on `this`, named for what it is (`ss`, `sheet`, `column`). `GenericSheetOperator extends SheetBaseNamed<SN>` with `ss`/`sheet`/`schema` getters is the reference shape — `sheet` is the primary (data) sheet, and the metadata view is `sheet.meta`; a column-scoped operator extends `ColumnBaseNamed<SN, CN>` and exposes a `column` getter the same way.

Inheriting the concrete class instead would put its whole surface on the operator, which is the opposite of what the operator is for — it exists to offer a *narrower*, more specific set of methods than the general class does.

## An Operator's props are its identity

What an Operator holds as props is what it *is*; a value that only one run cares about is an argument to the method that needs it. `OccupancyLedgerOperator` is constructed from the spreadsheet alone and its `build(occupancyRowIndex)` takes the occupancy per run, so any caller holding a spreadsheet can build a ledger and two builds in a row are independent.

Holding the row index as an optional field assigned at the start of a build was considered and rejected (#22). It gives the operator two lifetimes with nothing in the type separating them: before a build the field is unset and every step reading it fails, and after one returns the field still names the previous occupancy, so anything reading the operator then gets a confident answer about the wrong tenancy. That is [design.md](../design.md)'s "Make disagreement structurally impossible" applied to an operator's own state.

## Push a domain query onto the object that owns it

When a coordinating class composes several calls on a collaborator to answer one domain question, that composition belongs on the collaborator as its own named method — not re-inlined at every call site. `ColumnConfigOperator` used to reach through `sheet.uniformRow("columnId").activeValueArr` and `.hasValue(columnId)` directly; that logic moved onto `SheetMetaNamed` itself as `get activeColumnIds()` and `isActiveColumnId(columnId)`, and `ColumnConfigOperator`'s own private helper now just delegates:

```ts
private _isActiveColumnId(sheetGid: number, columnId: string): boolean {
  return this.ss.raw.sheetMeta(sheetGid).isActiveColumnId(columnId);
}
```

Destructure a collaborator's getter directly when only one property is needed: `const { activeColumnIds } = this.ss.raw.sheetMeta(sheetGid);`.

A container method that takes an index/id as a parameter, but is only ever called by code that already has that exact value as its own instance state, is a sign the query belongs on the instance instead — drop the parameter along with the method. `SheetBaseRaw.columnValidationValues(colIndex: number)` was deleted; its one caller always already had its own `colIndex`, so the query moved to `ColumnMetaRaw` as `get valueValidationStrings()`, reading `this.sheet.activeTable.columnValidationValues.get(this.colIndex)` (`ColumnMetaRaw.ts`). The parameter disappearing is what turns it into a getter (see the getter rule in `@byronbroughten/config`'s `docs/style/naming.md`).

## Model state at the granularity the concept actually has

The principle and its other instances live in [design.md](../design.md); what follows is where it lands on member placement. A member that **samples the top data row to derive a column-wide fact** belongs on the Meta column — that is membership criterion 2 of the Meta/primary axis ([vocabulary.md](../vocabulary.md), "Meta / primary"), and `isFormula`/`numberFormatType` are the case that produced it. They match `ColumnSchema.isFormula`, the schema-based trait, and are read off the column's top data-row cell only because that is how the API delivers them; so they live as `activeIsFormula`/`activeNumberFormatType` on `ColumnMetaRaw`, populated once per column by `SheetRaw._integrateSheetData`, not as per-row/per-cell state on `CellRaw`/`RowBaseRaw`.

The criterion bites on the derived fact, not on row or cell addressing: `topCell` and `topRow` stay primary, or a Meta class would end up handing out data rows.

