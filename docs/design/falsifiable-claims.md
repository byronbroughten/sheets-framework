# Make the claim falsifiable before believing it

Design reasoning. The one-line principle lives in [`docs/design.md`](../design.md); this file holds the argument and its instances.

## The argument

Intuition about cost and intuition about types are both unreliable here, and both have a cheap check available. Run the check, then record the number or the assertion so the next person doesn't re-litigate it.

## Instances: performance

`SpreadsheetApp` "should" be cheaper inside a trigger that already has the sheet open. Measured, it is slower — each call is its own round trip (`docs/architecture/round-trips.md`). Grouping every value name up front "should" beat re-deriving per use; measured, eager cost ~390k instantiations against ~48k per lazy use (`docs/architecture/type-check-cost.md`).

## Instances: types

An assignment proves nothing about a mapped or conditional type — it passes against `any` and against `never` alike. Identity-based `IsExactly`/`assertType` is the only probe that means anything, and a probe that needed an `any` to compile has proved nothing at all: intersecting to satisfy an indexer resolves to `any` and makes every downstream assertion vacuously true (the framework's docs/style.md, "Type modeling").

## Instances: live behaviour

The same holds for facts about Sheets that Google's docs don't state, where the cheap check is a probe against the live spreadsheet. An edit-protection probe recorded that an open-ended column protection did not cover a Table row added after it was set; a follow-up check found that it does, and the shape decision had briefly been justified by the range merely *reading back* open-ended — a proxy for coverage rather than coverage itself, which is how a decision order written around the proxy produced the wrong branch. The same write-up recorded that Apps Script's "the owner can never be removed as an editor" fails for a REST write, on one observation with an ordinary alternative explanation available: a dotted gmail address may name a different Google identity than the undotted one, in which case the owner was simply not on the list. That claim is withdrawn and the question left open rather than answered (#51, #49; `docs/architecture/edit-protections.md`).

## Corollaries

Record the measurement next to the conclusion. "`SpreadsheetApp` is slower" is re-proposable; "measured ~494ms against ~350ms on this date" is not.
