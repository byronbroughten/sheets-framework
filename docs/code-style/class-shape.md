# Class shape

Style fragment. The one-line rules live in [`docs/code-style.md`](../code-style.md); this file holds the reasoning and the worked examples. The general class-shape reasoning (member order, helpers, dead scaffolding) is in `@byronbroughten/config`'s `docs/code-style/class-shape.md`.

One heading per class-shape rule, most of them with the refactor that produced it: grep `^## ` for the rule you're applying and read that section.

## Coordinator classes

When code coordinates other stateful objects (other Operators, a Spreadsheet), write it as a **coordinator class** extending the tier's Base class (matching `TableConfigOperator`).

This was mined from the framework tiers. It does **not** govern business endpoints: an endpoint is a plain entry with module-private helpers, matching the endpoints already in the registry. Reach for a class only once a file is unwieldy or its logic finds a second caller.

**What unwieldy meant, the one time it was crossed.** `buildLedger.ts` was 303 lines holding eighteen free functions, and almost every one of them existed to thread a collaborator through: the spreadsheet appeared in six of their signatures, the occupancy id in three, and a map of charges keyed by id in one more (#22). The number is the threshold's first data point, not its definition — what made it unwieldy was the threading, and the win was signatures rather than lines, since the body came out about the same length as `OccupancyLedgerOperator`.

- **`init`** is how a caller outside the class builds one; **`new`** is how a class builds its own collaborators from props already on `this`. `init` takes whatever the caller already holds — nothing at all for an entry point that starts a run (`ConfigCoordinator.init()`), or a live collaborator whose state must be shared, as when an endpoint action builds an operator from the `ss` it was handed — and assembles the props itself. A collaborator reached from props already on `this` skips `init` and is constructed directly in a getter: `new ColumnConfigOperator(this.operatorProps)` when the collaborator needs config-sync state, otherwise `new SpreadsheetNamed(this.spreadsheetNamedProps)`. Either way the real constructor just takes a `props` object.
- Collaborators (`ss`, `tableConfigOperator`, `schema`, etc.) are lazy getters built from shared props on `this`.

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
  get tableConfigOperator() {
    return new TableConfigOperator(this.operatorProps);
  }
  get columnConfigOperator() {
    return new ColumnConfigOperator(this.operatorProps);
  }
  syncAndFlushConfigSheets() {
    this.tableConfigOperator.fetchAndUpdateAll();
    this.columnConfigOperator.fetchAndUpdateColumnConfig();
    this.ss.batchUpdateGSheets();
  }
}
```

Callers reach `orchestrator.tableConfigOperator` on the instance. The method returns nothing.

## A coordinator splits by the general rule

When, how and why a class splits into collaborators is `@byronbroughten/config`'s `docs/code-style/class-shape.md`, "Split a class when its helpers fall into groups that share nothing" and the section after it. `SpreadsheetRaw` is the framework's instance: 26 private helpers in three groups, finishing a gathered fetch, checking that each sheet's Table is where the layout requires, and turning queued changes into a batch update and sending it (#67). What the framework adds:

- **Collaborators are built the coordinator way above**, with `new` from the coordinator's props in lazy getters (`fetcher`, `flusher`); `SpreadsheetFetcherRaw` builds its own `tableValidator`, because finishing a fetch is what decides which Tables to judge.
- **Collaborators get no tests of their own**: the coordinator keeps each public method as a one-line delegation, so its tests already cover them.
- **A base class stays in `ClassBases/`**, never in one coordinator's subfolder (`02_SpreadsheetRaw/SpreadsheetRaw/`), because it serves a whole chain of classes, often across tiers.
- **The getter back to the coordinator is named for its shared surface**: `ss` for a spreadsheet coordinator, as `SpreadsheetFlusherRaw` does, and `sheet` for a sheet one, as `SheetEditProtectionsRaw` does, rather than copying `sheet(sheetGid)` into each collaborator (byro-repo#344).

## An Operator extends a `*BaseNamed` and reaches its subject through a getter

Whatever an Operator operates on — a Table, a column — it extends that thing's `*BaseNamed` class and adds methods suited to that data structure. It does **not** extend the concrete class it works through, and it doesn't take one as a constructor argument: the subject is a lazy collaborator getter built from the props already on `this`, named for what it is (`ss`, `table`, `column`). `GenericTableOperator extends TableBaseNamed<TN>` with `ss`/`table`/`schema` getters is the reference shape; a column-scoped operator extends `ColumnBaseNamed<TN, CN>` and exposes a `column` getter the same way.

Inheriting the concrete class instead would put its whole surface on the operator, which is the opposite of what the operator is for — it exists to offer a *narrower*, more specific set of methods than the general class does.

## An Operator's props are its identity

What an Operator holds as props is what it *is*; a value that only one run cares about is an argument to the method that needs it. `OccupancyLedgerOperator` is constructed from the spreadsheet alone and its `build(occupancyRowIndex)` takes the occupancy per run, so any caller holding a spreadsheet can build a ledger and two builds in a row are independent.

Holding the row index as an optional field assigned at the start of a build was considered and rejected (#22). It gives the operator two lifetimes with nothing in the type separating them: before a build the field is unset and every step reading it fails, and after one returns the field still names the previous occupancy, so anything reading the operator then gets a confident answer about the wrong tenancy. That is [design.md](../design.md)'s "Make disagreement structurally impossible" applied to an operator's own state.

## A query whose parameter its caller already holds belongs on the instance

This is the framework's case of `@byronbroughten/config`'s rule that a helper on one collaborator's state belongs on that collaborator (`docs/code-style/class-shape.md`, "A helper that never reads `this` is a module function"). `ColumnConfigOperator` used to reach through `sheet.headRow("columnId").workingValueArr` and `.hasValue(columnId)` directly; that logic moved onto the Table profile as `get columnIds()` (`TableProfileRaw`) and onto the Table's column resolver as `hasColumnId(columnId)`, and `ColumnConfigOperator`'s own private helper now just delegates:

```ts
private _hasColumnId(tableId: string, columnId: string): boolean {
  return this.ss.raw.table(tableId).columnResolver.hasColumnId(columnId);
}
```

Destructure a collaborator's getter directly when only one property is needed: `const { columnIds } = this.ss.raw.table(tableId).profile;`.

A container method that takes an index/id as a parameter, but is only ever called by code that already has that exact value as its own instance state, belongs on the instance instead — drop the parameter along with the method. `SheetBaseRaw.columnValidationValues(colIndex: number)` was deleted; its one caller always already had its own `colIndex`, so the query moved onto the column as `get valueValidationStrings()`, now on its profile (`ColumnProfileRaw.ts`). The parameter disappearing is what turns it into a getter (see the getter rule in `@byronbroughten/config`'s `docs/code-style/naming.md`).

## Model state at the granularity the concept actually has

The principle and its other instances live in [design.md](../design.md); what follows is where it lands on member placement. A member that **samples the top data row to derive a column-wide fact** belongs on the column profile, and `isFormula`/`numberFormatType` are the case that produced it. They match `ColumnSchema.isFormula`, the schema-based trait, and are read off the column's top data-row cell only because that is how the API delivers them; so they are held once per column in `ColumnStateRaw.sampledFacts`, filled by `TableRaw`'s fetch integration and read as `column.profile.isFormula`/`.numberFormatType`, not as per-row/per-cell state on `CellRaw`/`RowBaseRaw`.

The criterion bites on the derived fact, not on row or cell addressing: `topCell` and `topRow` stay on the column and Table, or the profile would end up handing out data rows.

