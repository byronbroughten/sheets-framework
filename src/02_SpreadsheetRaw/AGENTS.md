# Rules for `src/02_SpreadsheetRaw/`

- **A new write is queued and gathered on the Table, row or cell; the flusher only sends the batch.**
- **A new operation kind settles four things**: its queue key, the state that holds it, the gather method that builds it, and its slot in the flush order.
- **A new write method whose request embeds a row coordinate calls `this.table.assertRowIndexesNotStale()` (on the Table itself, `this.assertRowIndexesNotStale()`) on its first line**, in the same change; a whole-column protection is exempt.
- **A gathered write converts its rows through `originAtGathering()`**, which holds the stale check for every gathered operation.
- Mechanics: [`docs/architecture/queued-writes.md`](../../docs/architecture/queued-writes.md) and [`working-view.md`](../../docs/architecture/working-view.md).
