# Make the claim falsifiable before believing it

Design reasoning. The one-line principle lives in [`docs/design.md`](../design.md); this file holds its instances.

## Instances: performance

[ADR 0025](../adr/0025-read-the-sheet-through-the-sheets-api-not-spreadsheetapp-even-in-a-trigger.md) keeps a trigger's reads off `SpreadsheetApp`, which "should" be cheaper with the sheet already open and measured slower. [ADR 0004](../adr/0004-filter-the-column-map-per-use-rather-than-group-it-up-front.md) filters the column map per use, where grouping every value name up front "should" have been cheaper and measured about eight times dearer. [ADR 0007](../adr/0007-key-an-endpoint-by-the-column-that-fires-it-and-tie-its-columns-to-that-sheet.md) measured its per-key sheet correlation before choosing it, and records the number. [ADR 0010](../adr/0010-type-a-fixed-list-dropdown-column-as-its-literal-union.md) records beside its decision what the `const` type parameter it rests on costs. [ADR 0015](../adr/0015-leave-relative-and-absolute-column-addressing-unbridged-at-the-type-level.md) rejected a type-level bridge between the two column-addressing families on a prototype's measured `TS2590` and check time, not on a guess.

## Instances: types

For a type, the cheap check is an identity assertion rather than an assignment, and a probe that needed an `any` to compile has proved nothing; both, with the indexer case that taught the second, live in [`docs/code-style/type-modeling.md`](../code-style/type-modeling.md#identity-checks-never-assignment).

## Instances: live behaviour

The same holds for facts about Sheets that Google's docs don't state, where the cheap check is a probe against the live spreadsheet. An edit-protection probe recorded that an open-ended column protection did not cover a Table row added after it was set; a follow-up check found that it does, and the shape decision had briefly been justified by the range merely *reading back* open-ended — a proxy for coverage rather than coverage itself, which is how a decision order written around the proxy produced the wrong branch. The same write-up recorded that Apps Script's "the owner can never be removed as an editor" fails for a REST write, on one observation with an ordinary alternative explanation available: a dotted gmail address may name a different Google identity than the undotted one, in which case the owner was simply not on the list. That claim is withdrawn and the question left open rather than answered (#51, #49; `docs/architecture/edit-protections.md`). [ADR 0002](../adr/0002-protect-the-config-sheet-floor-with-edit-warnings-not-edit-locks.md) rests the floor's warnings-only shape on that probe's observation that a lock naming no editors stops nobody. Table growth was settled the same way: live probes on the dev spreadsheet established what `insertRange`, `deleteRange`, `appendDimension` and `copyPaste` do to a Table and its neighbours before the growth recipe relied on them, and the recipe's cost was measured rather than guessed: growing by 1,500–5,000 rows in one batch cost 8–18% more than `appendCells`, before the format copy (sheets-framework#53, sheets-framework#56, sheets-framework#59, sheets-framework#61; `docs/testing.md`). [ADR 0013](../adr/0013-grow-a-table-by-insert-plus-widen.md) is the recipe those probes cleared, chosen over `appendCells` for failure modes the probes observed.

## Corollaries

Record the measurement next to the conclusion. "`SpreadsheetApp` is slower" is re-proposable; "measured ~494ms against ~350ms on this date" is not.
