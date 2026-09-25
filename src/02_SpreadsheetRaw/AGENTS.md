# Rules for `src/02_SpreadsheetRaw/`

- **A new write is queued and gathered on the sheet, row or cell; the flusher only sends the batch.** Row deletes, built in `SpreadsheetFlusherRaw`, are the one exception.
- **A new request kind settles four things**: its queue op, the state that holds it, the gather method that builds it, and its slot in the flush order.
- **A new write method whose request embeds a row coordinate calls `this.sheet.activeTable.assertRowIndexesNotStale()` on its first line**, in the same change; a whole-column protection calls `validateColIndexNotStale` instead.
- Mechanics: [`docs/architecture/queued-writes.md`](../../docs/architecture/queued-writes.md) and [`working-view.md`](../../docs/architecture/working-view.md).
