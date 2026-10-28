# Architecture

Map fragments, one file per heading. Open the file the task needs.

| When | File |
| --- | --- |
| Chore homes, no registry, not tested | [chores.md](./architecture/chores.md) |
| `gatherRawOperation` | [raw-request-opening.md](./architecture/raw-request-opening.md) |
| Conditional format rules (order, prepend, identity, anchored formulas) | [conditional-format-rules.md](./architecture/conditional-format-rules.md) |
| Edit protections (edit warnings, edit locks, identity, stale refetch) | [edit-protections.md](./architecture/edit-protections.md) |
| Queuing and gathering a write, the shared queue, discarding it | [queued-writes.md](./architecture/queued-writes.md) |
| The flush's send order, add-sheet and add-Table leading it | [flush-order.md](./architecture/flush-order.md) |
| Stale row indexes after a flush, the assert a new write ships | [stale-row-indexes.md](./architecture/stale-row-indexes.md) |
| Row deletes and sorts bounded to the Table | [row-deletes-and-sorts.md](./architecture/row-deletes-and-sorts.md) |
| Growth: `appendTableRows`, its model row, growing from a blank row | [table-growth.md](./architecture/table-growth.md) |
| Column inserts, `updateColumnType` limits | [table-columns.md](./architecture/table-columns.md) |
| Stacked Tables shifting in one batch, split and partial-Table refusals | [stacked-tables.md](./architecture/stacked-tables.md) |
| Per-cell writes, column fills, formula writes, which queued write wins | [cell-writes.md](./architecture/cell-writes.md) |
| `findReplace`: fields, scope, the reads it invalidates | [find-replace.md](./architecture/find-replace.md) |
| Working vs fetched, the working view, writes without a fetch, reads that need one, what Google omits from a reply | [working-view.md](./architecture/working-view.md) |
| A missing Table, or one outside the header zone | [table-placement.md](./architecture/table-placement.md) |
| Routes for finding a moved Table, probed live: named ranges, developer metadata, whole-sheet `tables`, zone depth, and what each costs | [finding-tables.md](./architecture/finding-tables.md) |
| Last data row, a header-only Table, blank-row reuse | [blank-row.md](./architecture/blank-row.md) |
| `appendRowWithVals` vs `appendRowWithAllVals`, what the required bag doesn't filter | [row-appends.md](./architecture/row-appends.md) |
| Endpoint entry, `EndpointRun`, run report, selector behavior | [endpoint-dispatch.md](./architecture/endpoint-dispatch.md) |
| Sheets round trips | [round-trips.md](./architecture/round-trips.md) |
| Instantiation budget, the template-literal cliff | [type-check-cost.md](./architecture/type-check-cost.md) |
| Relative `<TN, CN>` vs `ColumnFullName` | [column-addressing.md](./architecture/column-addressing.md) |
| `SpreadsheetBaseSchema` / `SpreadsheetSchema` / `TableSchema` / `ColumnSchema` | [schema-classes.md](./architecture/schema-classes.md) |
| Table and column class chains, `TableCommonRaw`, the Sheet container | [class-chains.md](./architecture/class-chains.md) |
