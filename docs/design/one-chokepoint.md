# Funnel the expensive thing through one place

Design reasoning. The one-line principle lives in [`docs/design.md`](../design.md); this file holds the argument and its instances.

## The argument

Design so the costly operation has exactly one chokepoint. Then instrumenting it measures everything, optimizing it optimizes everything, and a new call site can't quietly add cost behind your back.

## Instances

Every Sheets read goes through `fetchAllGathered` and every write through `_sendWriteOperations` — instrument those two and the whole system is measured (`docs/architecture/round-trips.md`). Writes don't hit the API at all until a flush: `update`/`append`/`delete` mutate local state and register a coordinate, and one `batchUpdate` ships the lot (`docs/architecture/queued-writes.md`). Because that chokepoint already existed, adding background-colour writes cost no new round trip — the colour merges into the update already queued for that cell (`b3eb76d`). The same gather is why N appended rows become one `appendTableRows` per Table rather than N: the queue groups them by `tableId`, and gathering expands each group into one insert, widen and format copy sized for all its rows (`docs/architecture/queued-writes.md`). The Node host applies the same argument one level below the framework: its adapter is the single door to Google, so arming a dry run there makes a *writing* dry run unrepresentable rather than discouraged — no flush anywhere, the config orchestrator's internal one included, can reach the sheet, and no caller has to be trusted to withhold one. Narrowing a chore's spreadsheet handle to remove the flush was considered and dropped as buying nothing the funnel doesn't already give (#25).

## Corollaries

The chokepoint is also where shared state bites. One `spreadsheetStateRaw` threaded by reference means a flush commits everything queued anywhere in the run — which is why a failure path must discard before it writes status, or the `finally` ships a half-finished run alongside its own error report.
