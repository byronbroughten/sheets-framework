# Queued writes and the flush

Map fragment. Sibling headings live in this folder. What a run sees between fetch and flush: [the working view](./working-view.md).

A write mutates local state and queues an operation; one `batchUpdate` per flush sends them, in the [flush order](./flush-order.md). The rules for adding a write: [`src/02_SpreadsheetRaw/AGENTS.md`](../../src/02_SpreadsheetRaw/AGENTS.md).

## Writes queue locally until a flush

`update`/`append`/`delete` don't hit the API — they mutate the working view and register the write on that sheet's `writeQueue`. Nothing reaches the spreadsheet until `batchUpdateGSheets()` turns each sheet's queued writes into write operations and `RawSource.flush` sends them. `GoogleSheetsAPI` maps each local operation onto Google request kinds (`updateCells`, `pasteData`, `repeatCell`, `insertRange`, `updateTable`, …) and one `batchUpdate`.

## A new write is queued and gathered on the Table, row or cell

A new write is queued and gathered on the Table, row or cell, and the flusher only sends the batch. The Table, sheet, row or cell builds the operation, either when the write is queued (`SheetRaw.updateTitle`) or from the write queue at flush time (`TableRaw.gatherFillColumnOperation`). A new operation kind therefore settles four things: its queue key, the state that holds it, the gather method that builds it (none for a spreadsheet-level op like `addSheet`, since a tab that doesn't exist yet has no sheet state to gather from), and its slot in the [flush order](./flush-order.md).

## One queue shared by reference, and discarding it

The run's `spreadsheetStateRaw` is threaded **by reference** through `spreadsheetNamedProps` into every collaborator getter, so an endpoint, its `ConfigCoordinator`, and every operator beneath it all share one queue. A flush therefore commits everything queued anywhere in the run, not just what the calling object queued — which is why a failure path must `discardQueuedChanges()` before writing status, or the `finally` flush ships a half-finished sync. Discarding resets both write queues as groups (the spreadsheet's `writeQueue.operations` and every sheet's `writeQueue`) rather than re-initializing individual fields. It leaves the working view reflecting writes that never happened, so it's a terminal step only.
