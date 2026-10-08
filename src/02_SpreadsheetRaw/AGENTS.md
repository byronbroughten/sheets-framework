# Rules for `src/02_SpreadsheetRaw/`

- **A new write is queued and gathered on the Table, row or cell; the flusher only sends the batch.**
- **A new operation kind settles four things**: its queue key, the state that holds it, the gather method that builds it, and its slot in the flush order.
- **A new write method whose request embeds a row coordinate calls `this.table.assertRowIndexesNotStale()` (on the Table itself, `this.assertRowIndexesNotStale()`; on a sheet, `this.sheet.assertRowIndexesNotStale()`) on its first line**, in the same change; a whole-column protection is exempt.
- **A gathered write converts its rows through `originAtGathering()`**, which holds the stale check for every gathered operation.
- **A write reached through a Table's row, column or cell queues a Table-relative range and converts it at gathering**, so a same-batch shift carries it; checkboxes, conditional-format adds and protections included. Only a whole-sheet protection is built when queued.
- Mechanics: [`docs/architecture/queued-writes.md`](../../docs/architecture/queued-writes.md) and [`working-view.md`](../../docs/architecture/working-view.md).
