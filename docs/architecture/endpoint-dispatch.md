# Endpoint dispatch

Map fragment. Sibling headings live in this folder.

**What an endpoint, a runner, a two-way endpoint, a selector, a run status, a run state and a run report *are* is defined in [`CONTEXT.md`](../../CONTEXT.md)** — read that first; this file is only how the dispatch is built.

An endpoint is one entry keyed by the column whose action-row checkbox triggers it, and `EndpointRun` owns its run: selection, running state, action, outcome, feedback. The rules for writing one: [`src/06_API/AGENTS.md`](../../src/06_API/AGENTS.md), plus the app's own endpoint-folder rules.

## The endpoint entry

An endpoint is **one entry keyed by the column whose action-row checkbox triggers it**. Any column may be that key — there is no suffix requirement and no second endpoint kind. The entry's value names an action plus, optionally, the columns the framework manages on the endpoint's behalf — two feedback columns, and a selector declared as an object so the opt-out from clearing sits inside the thing it modifies:

```ts
export interface Endpoint<SN extends SheetNameSimple> {
  action: EndpointAction;
  timeLastRan?: FeedbackColumnName<SN>;
  runStatus?: FeedbackColumnName<SN>;
  selector?: {
    column: CheckboxColumnName<SN>;
    retainSelection?: boolean;
    requireOneRow?: boolean;
  };
  runOnUncheck?: boolean;
}

export type EndpointsAll = {
  [FN in ColumnFullName]?: Endpoint<SheetNameOf<FN>>;
};

export type Endpoints = {
  [FN in ColumnFullName as SheetNameOf<FN> extends FloorTabName
    ? never
    : FN]?: Endpoint<SheetNameOf<FN>>;
};
```

`Endpoints` is the app's map: a key on one of the four config sheets is a type error, because the framework owns those entries (`frameworkEndpoints`). `Api` spreads the framework's entries last into an `EndpointsAll`, so a collision still resolves to the framework at runtime.

`action` is the only required field, so an endpoint takes exactly the machinery it wants: a bulk "select all" declares an action alone and stays at the round-trip floor; a run over a selection declares the feedback columns and the selector too. `retainSelection` and `requireOneRow` are nameable only inside a declared selector, so an endpoint without one cannot carry either as a silent no-op. Register endpoints with a plain `: Endpoints` annotation — **not** `makeStructuredConfig`, which lets an unknown key through (the house style's "Type modeling").

## Flags are directives to the framework

**Every flag reads as a directive to the framework** — `retainSelection`, `requireOneRow`, `runOnUncheck` — rather than as a description of the entry it sits on (#18). One mood for the whole entry means a reader never has to work out which flags command and which describe.

## One endpoint, one file

**One endpoint, one file.** `businessEndpoints.ts` stays a registry: each entry's body lives in `businessEndpoints/<name>.ts` as a single exported `Endpoint<"thatSheet">` the registry imports and assigns to its key, so an endpoint's size never crowds its neighbours and the plain `: Endpoints` annotation still catches a key naming a column that doesn't exist. `businessEndpoints/buildLedger.ts` is the reference; `businessEndpoints/updateTerms.ts` was moved out to its own file once the pattern had been proven on a new endpoint, its body carried over whole. `businessEndpoints/addPropertyExpense.ts` was the last inline entry — an empty action with nothing to move — and arrived as a file of its own with the work behind it (#24), so no entry is inline any more.

## Two shapes for an endpoint body

**An endpoint body takes one of two shapes.** The default is module-private helpers around the entry, which `businessEndpoints/updateTerms.ts` is the reference for. Once a body turns unwieldy the logic moves onto a business operator under `businessEndpoints/BusinessOperators/`, and the entry keeps a one-line action delegating to it — `businessEndpoints/buildLedger.ts` is the reference for that shape, and its operator is `OccupancyLedgerOperator` (#22); `businessEndpoints/addPropertyExpense.ts` is the second, behind `PropertyExpenseOperator` (#24). The threshold that decides between them lives in the house style's "Coordinator classes", which was mined from the framework tiers and otherwise doesn't govern endpoints.

## No operator scoped to the occupancy row

**An operator scoped to the occupancy row is deliberately absent.** Building a ledger reads exactly three cells off that row — its id, its name, and its ledger start date — and three plain reads are not a domain query — such a class would be created empty and do nothing a getter does not. The trigger for building one is #21: the update-terms body reads twenty-odd next-terms columns, validates a start date against the latest terms, closes the open term and appends a new one, which would give the class real behaviour and a second caller on the same day.

## Each key is correlated with its own sheet

**Each key is correlated with its own sheet**, so `SheetNameOf<FN>` narrows every column the entry names to that sheet's columns, and each is filtered to the value type it needs. A column from another sheet, a boolean column in `runStatus`, a string column in the selector's `column`, and a key that is not a column are each a compile error rather than a runtime no-op. It works because recovering a sheet name *from* a column full name is a plain indexed lookup on the flat column map — the cheap direction (see [Two ways to address a column](./column-addressing.md)); the expensive direction is never used. Measured: the map itself costs ~200 instantiations and no measurable check time.

## The steps of a run

**`EndpointRun` owns the run; `Api` only decodes and dispatches.** In order, a run:

1. preps the selector column's fetch, if one is declared, into the cycle `fetchAllPrepped` is already running — so a selection costs no round trip of its own, and an endpoint without one costs no read at all here;
2. collects the checked row indexes, or — with no selector — every table data row index from the table bounds;
3. resets the entry column's action cell, unless `runOnUncheck` makes the checkbox an input rather than a button;
4. **prunes**, when a selector is declared: every unselected data row is removed from local state, so every later read of active rows means the selection without a call site being rewritten. `SheetRaw.removeRowsExcept` keeps the uniform rows (dropping the columnId row would break column resolution), and marks the sheet, so a whole-column fill on it throws;
5. stamps the running state and flushes — that first flush is what puts "Running…" and yellow on the sheet *before* the work starts, which is the whole basis for a killed run staying distinguishable from one that never began;
6. **refuses a selection of more than one row**, when the selector declares `requireOneRow` — as the *first statement inside the `try`*, so the operator sees the ordinary failed run state carrying the count. That placement is the rule spelled out in step 9: a precondition that must fail loudly goes inside the `try` and before the action, since throwing earlier would write no status at all and the error path's discard would take the entry checkbox's reset with it;
7. runs the action inside `try`/`catch`/`finally`;
8. **clears the selection**, when one is declared and the endpoint doesn't retain it: the selected rows' selector cells are filled with an explicit `false`. It sits inside the `try`, after the action returns and before the outcome is written, so "clears on success" falls out of that placement rather than out of a flag — the error path discards the run's queued changes, which drops this fill along with everything else the action queued;
9. writes the outcome — whatever the action's **run report** says, or `Succeeded` when it says nothing — and flushes again. A failure `discardQueuedChanges()` first, or the `finally` ships a half-finished run alongside its own error report. That discard covers everything the `try` body queued, the action's own writes and the framework's alike — so any *other* step that must happen only on success needs no success flag either, just a position after the action and inside the `try`. The mirror holds for a precondition that must fail *loudly*: it belongs inside the `try` **before** the action. Throwing earlier writes no status at all, and the error path's discard would take the entry checkbox's reset with it, leaving a run button stuck ticked.

## What the operator sees of a selector

Ticking rows picks them out; the run then acts on those rows and reports into those rows, and leaves every other row alone. A successful run **consumes** its selection — the ticks clear themselves, the way the run button does, so an empty selector column means nothing is selected and the next run costs what it looks like it costs. A run that fails leaves the ticks alone: they are the operator's input, and the same selection can be retried once the problem is fixed.

An endpoint whose selector marks a standing set of rows rather than a one-off pick declares that it **retains its selection**, and its ticks survive a successful run untouched. An endpoint whose work is about one row and could not be about two — a ledger is one page about one tenancy — declares that it **requires one row**, and a run with more than one ticked fails before it starts, saying how many you ticked, leaving every tick where it is so you can untick the extras and go again.

## The run report

**An action returns a run report**, which is nothing, a string, or an object carrying a run state, a message and a map of per-row reports — every part optional:

```ts
type RunReport =
  | { runState?: "success"; message?: string }
  | { runState: "warning" | "failure"; message: string };

type ActionReturn = void | string | (RunReport & { rows?: Map<number, RunReport> });
```

A bare string still means success with that message, and nothing at all still means a bare `Succeeded`, so the two shapes that existed before this arrived are untouched. Warning and failure demand a message, because neither has a sentence worth writing on its own. The run has four states — `running`, `success`, `warning`, `failure` — each paired with its colour in the one `runStates` object, so no path can show one state's colour beside another's message.

A **warning**'s orange sits between the success green and the failure red, so the three read as a scale, and it always carries a sentence of its own, since an orange cell with nothing to say would be a puzzle.

## Rows the report doesn't name take the top-level state

**The top-level state is the default for rows the map doesn't name, not a claim about the run.** The framework derives nothing from the map: there is nowhere run-wide to display a state, since both feedback columns are columns and every state an operator sees is already per row. That also removes the case where an action contradicts itself between the two.

## Per-row report keys are grid row indexes

**The map is keyed by the base-zero grid row index the action was handed as its selection**, and a key that is not one of the sheet's data rows throws — a reporting bug is loud rather than writing somewhere surprising. It throws from inside the run's `try`, so the operator reads it as an ordinary failed run. A key naming a row the same action deleted is legal and costs no check: the delete wins within the batch either way.

## A returned failure is not a thrown one

**A returned failure is not a thrown one.** A throw abandons the run, discards everything queued and paints every row red; a returned failure means that one row did not go through while the rest of the run committed. Both are red, because the operator acts on the same fact — this row did not convert — and the sheet already distinguishes them: after a throw every row is still there and every row is red, after a partial run only the refused rows are.

## A selector with nothing ticked skips the action

An endpoint with a selector but nothing ticked prunes every data row, leaving no cell to report into: it skips the action, logs, and flushes only the checkbox reset. Running an action against an empty selection would invite domain code that reads "nothing selected" as "everything".

## How feedback is written

Feedback is written with the two column fills from [column fills](./queued-writes.md#column-fills), **both feedback columns taking the run state's colour**: `updateActiveCells` when there's a selector (the selected rows are active by construction), `updateAllCells` when there isn't — that branch reaches rows nothing ever fetched, which is why whole-sheet feedback costs no read. The selection clearing uses the active-cells fill for the same reason, through `CheckboxColumnOperator.uncheckActiveCells` — the operator's only uncheck, since a whole-column one refuses to run on a pruned sheet by design and would throw in exactly the situation that wants it. Contiguous selected rows collapse into one request each, and the fill rides the flush the `finally` was already going to perform, so the clearing adds no round trip.

## A per-row report is a per-cell write

**A per-row report is a per-cell write on top of that fill, and costs no round trip either.** `_sendUpdateRequests` already sends fills before per-cell updates and per-cell updates before row deletes, so the ordering the reporting needs falls out of the queue it was built on rather than out of a new rule. A per-cell write carries a field mask naming only the fields it sets, so a per-row report writes its own colour explicitly instead of leaning on the fill it overwrites; `CellIdentified.update` takes value and colour together for that reason, and the two halves merge into the one queued change.

## Clearing the rows you report into erases the message

**An action that clears or deletes the rows it reports into erases its own success message.** `clearValues` writes `""` into every non-formula column, the run-status column included, and per-cell updates land *after* the run's status fill within a flush — so a run that empties its sheet ends with an empty, green-tinted status cell rather than `Succeeded`. That is usually the right outcome, since the rows disappearing is the feedback, but such an endpoint can only report on the rows it leaves behind. The add-property-expense endpoint is built on exactly that: it deletes every row it converts, so a clean batch ends blank and a partial one leaves only the refused rows, each carrying its own message (#24). A failure is unaffected: the error path discards the clearing along with everything else the action queued.

## `EndpointRun` is built from `Api`'s props

`EndpointRun` is constructed from `Api`'s `SpreadsheetNamedProps` (the rule: [`src/06_API/AGENTS.md`](../../src/06_API/AGENTS.md)), never from a no-arg `init()` that mints fresh `spreadsheetStateRaw`. `Api` has already fetched the sheet's properties and columnId row by the time it dispatches; minting fresh state re-fetches all of it.

## The dispatch boundary widens the generic

**The dispatch boundary is where the generic widens.** There is deliberately no type-level bridge from a column full name to a sheet-and-column pair, so `Api` — holding a full name resolved at runtime — instantiates `EndpointRun` at the widened sheet name, where a column parameter is the union across sheets rather than one sheet's. That is sound and does not collapse to `never`, because `ColumnNameFiltered` distributes over the sheet name; `Endpoints.test.ts` pins both ends.

That widening is what forces the selector's shape to be spelled inline, and the run to take `EndpointDispatched` (the rule: [`src/06_API/AGENTS.md`](../../src/06_API/AGENTS.md)). Two generic references to the *same* named type are compared by that type's measured variance rather than property by property, and the column filter leaves the variance unmeasurable, so the comparison falls back to demanding identical sheet names. Nesting the selector inside a named `EndpointSelector<SN>` — interface or alias — therefore breaks `Api`'s assignment outright, and so does `Endpoint<SheetNameSimple>` as the run's prop type; an anonymous nested object plus a structural copy (`{ [K in keyof Endpoint<SN>]: Endpoint<SN>[K] }`) keeps both comparisons structural. Tidying either into a named type fails `npm run tsc` at `Api.ts`, not at the file you edited.

## An endpoint that appends into its own sheet

An endpoint whose action appends into its **own** sheet is a design smell. With no selector the run stamps its feedback across that sheet's data rows, so the run reports into rows it is still creating. Initiate and report such a run from a sheet other than the one it writes into.
