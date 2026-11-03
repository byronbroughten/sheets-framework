# Prototype: declarations and the three-way plan, on one real-estate Table

Throwaway prototype for [sheets-framework#183](https://github.com/byronbroughten/sheets-framework/issues/183), part of the map [#178](https://github.com/byronbroughten/sheets-framework/issues/178). It raises the fidelity of the model, file-shape, property and gate tickets; it decides none of them. This branch is never merged.

- **Run it:** open [`declarations-prototype/declarationsPlan.prototype.html`](./declarations-prototype/declarationsPlan.prototype.html) in a browser. Free-play buttons, eleven walkthroughs, three choices of what counts as "last time", and the JSON an agent would read.
- **Logic only:** `node declarations-prototype/agent-reader-test/run.js declarations-prototype/declarationsPlan.prototype.html` runs every walkthrough through the pure module and prints each step's plan.
- **Agent-reader test:** [`agent-reader-test/replies.md`](./declarations-prototype/agent-reader-test/replies.md) holds the inputs and what three fresh agents did with them.

The Table is occPayment: its column IDs, headers and `opy` prefix are real; column types, formula text, typed-cell counts and dropdown styles are illustrative. Deletion isn't modelled.

## Where the session landed

The developer's calls, recorded as leanings for the tickets that own them:

1. **"Last time" is the generated configs, refreshed on every run, a stopped run included; a declared property in conflict is held back at its last agreed value.** A regen never overwrites a declared property's value while the sheet disagrees with the declarations. With that rule the generated configs already hold the last agreed value, so no separate record of last applied values is needed. The prototype's mode "Last generated configs, refreshed every run, conflicts held back", walkthrough "Refresh first, then edit".
2. **The generated configs carry column type, dropdown rule and formula text.** Today's `columnConfigs.ts` holds `valueName` and `isFormula` only, so there is no "last time" to compare those three against.
3. **Agents run `gen:configs` before editing the declarations.** A fresh picture turns the "both changed" conflict on a property the agent newly claims into a plain agent change: in case C below the agent would have seen Number and declared Currency knowingly, with no stop.
4. **Ways out of a stop.** An agent adopts (sets the declaration to the sheet's value) without asking when that doesn't conflict with the spec or task it's working on. Otherwise it asks the developer for permission to force. Dropping the claim (deleting the property) sends nothing too.
5. **A yes is needed only for:** a write that overwrites typed data, with the cells listed; and forcing over a person's edit, whether or not data changes.
6. **Rejected: a five-minute editing window** after each run. A stale picture can only cause a stop, never a silent overwrite (a declared property compares against its held-back value; a newly declared one against the last refresh, so a later person's edit reads as "both changed"). Every sending run already refreshes first, in-process, within seconds. A window can't police file edits, only sends, so it would force two runs where one does. Five minutes is short for declare, regen, fix `tsc`, check. A refresh-age line in the output ("last refresh 40 minutes ago; 2 properties changed since") is the cheap nudge if one is wanted; a per-spreadsheet lock over read-compare-send is the tool if concurrent senders ever matter.

## Findings

### The snapshot trap is real, and the hold-back rule closes it

If "last time" is a copy of the sheet at the last regen, one regen while a person's edit is unresolved erases the evidence: the next run reads the old declared value as an agent's change and reverts the person silently (walkthrough "The snapshot trap"). A refresh before editing *is* that regen, so refresh-first without the hold-back rule makes the trap fire every time. With it, the stray regen keeps the agreed value and the run still stops.

**Residual edge:** a branch where the property is undeclared regenerates and captures a person's value; merged into a branch that declares it, the merged generated configs hold the person's value and the trap fires. A separate record of last applied values survives that merge. It needs concurrent branches that disagree about whether a property is declared plus a person editing it in between; noted, not designed around.

### A property declared for the first time has no agreed value yet

With a record-only baseline, every new declaration stopped as "unknown", the agent's own change included. Falling back to the last generated configs fixes it, and under the chosen design that fallback is simply what the generated configs hold.

### The agent-reader test: two of three right, one forced

See [`replies.md`](./declarations-prototype/agent-reader-test/replies.md). The output shapes that follow:

- **Force must say it needs the developer's yes,** and shouldn't be a ready-to-paste command. Case C forced because the gate rule named only data cells; case A called the paste-ready force line "routine-looking".
- **Adopt isn't neutral.** It can turn a `=list[…]` reference into a hard-coded list (A), or quietly abandon the task (C). The output should say what adopting does to the declaration and to the task.
- **Name what would be lost.** "Overwrites 41 typed cells" sent the agent to a probe. The gate needs the typed-vs-formula read anyway, and the declarations name the target columns before the read, so the cells' `userEnteredValue` can ride on the run's one read: no extra call.
- **The unmeasured chip-style reset** (a column-list resend) left two agents unsure whether it counts against the gate: [#189](https://github.com/byronbroughten/sheets-framework/issues/189) settles it.
- **Who changed the sheet, and when,** was missing. The Sheets API gives no per-cell author; Drive revisions would cost calls for little. Dropped.

### ID prefixes sit inside every column and row ID

`dimensionIds` (`src/01_SpreadsheetSchema/dimensionIds.ts`) builds `c:<prefix>:<suffix>` and `r:<prefix>:<suffix>`, and `idPrefix` is sampled back from the column-ID row. So an ID-prefix rename rewrites every key a declarations file keyed by column ID would use, and its `findReplace` reaches the Payment ID columns of occPayAllocation (`c:opa:lypy6OR`) and occPayDesignation (`c:opi:ozecESY`). A new column keyed by header needs rekeying once its ID is minted, too. Options weighed: fix a prefix when its Table is created and drop the rename (leaning: the prefix stays a debugging hint, keys never move, the cross-Table `findReplace` goes away, and an agent reads which Table a prefix names from `tableConfigs`); key by Table ID plus suffix; drop prefixes. Stable keys matter more for agents than readable ones: `tsc` catches a typo.

### Corrections to the property research (#179)

- **`SINGLE(Table[Header])` reads this row's cell inside a Table's data-row formula** (developer). The research quoted "#This Row currently is not supported" and concluded per-row formulas need A1 references. With `SINGLE`, a declared formula is the same text in every row and names columns by header, so comparing it against the sheet is one string. Unverified: whether a header rename rewrites `Table[Header]` references.
- **A declared formula needs no re-fill on growth.** A row appended through `appendCells` takes its formulas from a row the Table already has (observed, `docs/architecture/blank-row.md`), and a row a person adds in the Sheets UI does too (developer). The research's "must be re-filled whenever the Table grows" is wrong.

### Cost

None of this slows a run much. The baseline is local file work; the typed-cell read folds into the run's one read; the only added cost is the refresh run before editing, a few seconds per task.
