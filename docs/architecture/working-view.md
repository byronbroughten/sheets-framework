# The working view: active, fetched and queued

Map fragment. Sibling headings live in this folder. How the queue is sent: [queued writes and the flush](./queued-writes.md).

The working view is the fetched sheet plus every queued write, with row indexes at their pre-flush positions; "active" means present in it. A write needs no fetch, but a read does, and a queued write survives a re-fetch in the same run.

## Three groups of state per tier

Each tier's state has three groups. The **working** view is what the run sees: the fetched sheet plus every queued write, with row indexes at their pre-flush positions. The **fetch queue** is what the next read will fetch, and the **write queue** is what the next flush will send.

- **Raw spreadsheet:** `fetchQueue.gridRanges` for the next `fetchAllGathered`, and `writeQueue.updateRequests`, the requests a flush sends. The update requests stay spreadsheet-wide because the flush orders them by request kind across sheets.
- **Raw sheet:** `working` holds the title, known Table, rows, column states, conditional formats and edit protections. `fetchQueue` holds the gather flags for conditional formats and edit protections, plus `toFinalize`, the rows, columns and cells a fetch has filled but not yet finalized. `writeQueue` holds `sheet` (fills, sorts, column inserts), `rows` (each row's queued append, delete and cell updates, keyed by row index) and the row indexes an append has already handed out.
- **Identified sheet:** a `fetchQueue` only: fetch targets addressed by column ID and tagged by `kind`, plus gather flags that it passes down to Raw when it gathers the fetch.

## A write does not require a fetch

**A write does not require the row to have been fetched.** `update` queues its request either way and mirrors the value into local cell state only when the row is active; `validateIsWritable` needs nothing but the sheet properties. Feedback can therefore be written to every data row of a column without reading one of them.

## A queued write outlives a re-fetch

**A queued write outlives a re-fetch in the same run.** Between a fetch and the flush, local state is the live sheet plus the queued writes: a row queued for delete stays gone however it is re-fetched or backfilled, and a cell with a queued value — its own update, else the most recently queued value fill that covers it — keeps that value once the row is fetched. A queued tab title, Table name or column type is applied again, in queue order, on top of the sheet properties a fetch integrates, so the last one queued wins. Formula writes stay out of local state, so a re-fetch still shows the old effective value. The flush clears the queue, so a fetch after it integrates the live sheet only.

## Active means present in the working view

The corollary is the trap: **"active" means present in the working view, with row indexes still at their pre-flush positions** — not "exists on the sheet". Queued appends and removes update that view immediately; the flush is what would shift live indexes, and until then a deleted row stays inactive at the index it had when it was queued. On the `triggerOnEdit` path only the columnId row is ever fetched, so *no data row is active* — anything working from `rowIndexesActive` writes nothing at all there. `rowIndexesFull`, derived from the table bounds, is what names every data row.

## A read requires a fetch

**A read does require it.** `value`/`valueOrEmpty` throw when the row was never fetched, so any decision that _branches_ on a cell's current value carries a prefetch prerequisite the equivalent write doesn't. Queue the fetch in the same cycle, or decide at the call site what an unfetched row means.

## Presence in a range is not presence in a payload

**Presence in a range and presence in a payload are different facts.** A row's membership in the table buys nothing about the response: Sheets omits every cell that holds no value, no formula and no explicit number format, and a row where that is true of *every* cell comes back as a grid-data block describing each column with no `rowData` at all — or, past the populated grid, as no block at all. Measured against `Add Property Expenses` on 10 September 2026, its single blank data row sits squarely inside the table range and still arrives with no cells (#17). That is exactly the state the [blank-row](./blank-row.md) policy leaves behind, so it is ordinary rather than exceptional. `SpreadsheetRaw`'s finalize pass exists to close that gap: after integration it backfills a cell for every row and column that was fetched in full, and seeds blank **active facts** for the table columns the payload described no cell for, so a column that was fetched and is simply empty is indistinguishable from one whose top cell was empty inside a returned row. A column nobody fetched still throws, because that is a programmer error rather than an empty sheet.
