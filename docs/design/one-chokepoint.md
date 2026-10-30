# Funnel the expensive thing through one place

Design reasoning. The one-line principle lives in [`docs/design.md`](../design.md); this file holds its instances.

## Instances

Every Sheets read goes through `fetchAllGathered` and every write through `_sendWriteOperations` — instrument those two and the whole system is measured (`docs/architecture/round-trips.md`; [ADR 0025](../adr/0025-read-the-sheet-through-the-sheets-api-not-spreadsheetapp-even-in-a-trigger.md) keeps a trigger's reads in that gathered request rather than in `SpreadsheetApp` calls that are each a round trip of their own). Writes don't hit the API at all until a flush: `update`/`append`/`delete` mutate local state and register a coordinate, and one `batchUpdate` ships the lot (`docs/architecture/queued-writes.md`). Because that chokepoint already existed, adding background-colour writes cost no new round trip — the colour merges into the update already queued for that cell (`b3eb76d`). The same gather is why growth is one `appendTableRows` per Table rather than one request per row: the queue merges a Table's appended rows by `tableId`, and gathering expands each into the sheet's one `appendDimension`, an insert, a widen and a format copy sized for all its rows, followed by the value and formula fills ([ADR 0013](../adr/0013-grow-a-table-by-insert-plus-widen.md) queues each Table's growth once and expands it at gathering; `docs/architecture/table-growth.md`). The Node host applies the same argument one level below the framework ([ADR 0032](../adr/0032-enforce-a-chores-dry-run-at-the-node-hosts-adapter.md) arms the dry run at the adapter, the single door to Google, so no flush anywhere can reach the sheet and no caller has to be trusted to withhold one).

## Corollaries

The chokepoint is also where shared state bites. One `spreadsheetStateRaw` threaded by reference means a flush commits everything queued anywhere in the run — which is why a failure path must discard before it writes status, or the `finally` ships a half-finished run alongside its own error report.
