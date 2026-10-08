# Architecture

Map fragments, one file per heading. Open the file the task needs.

| When | File |
| --- | --- |
| Chore homes, no registry, not tested | [chores.md](./architecture/chores.md) |
| `gatherRawOperation` | [raw-request-opening.md](./architecture/raw-request-opening.md) |
| Conditional format rules (order, prepend, identity, anchored formulas) | [conditional-format-rules.md](./architecture/conditional-format-rules.md) |
| Edit protections (edit warnings, edit locks, identity, stale refetch) | [edit-protections.md](./architecture/edit-protections.md) |
| Queuing a write, flush order, stale indexes, row deletes, fills, formula writes, `findReplace`, discarding | [queued-writes.md](./architecture/queued-writes.md) |
| Working vs fetched, the working view, writes without a fetch, reads that need one | [working-view.md](./architecture/working-view.md) |
| A missing Table, or one outside the header zone | [table-placement.md](./architecture/table-placement.md) |
| Routes for finding a moved Table, probed live: named ranges, developer metadata, whole-sheet `tables`, zone depth, and what each costs | [finding-tables.md](./architecture/finding-tables.md) |
| Last data row, blank-row reuse | [blank-row.md](./architecture/blank-row.md) |
| Endpoint entry, `EndpointRun`, run report, selector behavior | [endpoint-dispatch.md](./architecture/endpoint-dispatch.md) |
| Sheets round trips | [round-trips.md](./architecture/round-trips.md) |
| Instantiation budget, the template-literal cliff | [type-check-cost.md](./architecture/type-check-cost.md) |
| Relative `<TN, CN>` vs `ColumnFullName` | [column-addressing.md](./architecture/column-addressing.md) |
| `SpreadsheetBaseSchema` / `SpreadsheetSchema` / `TableSchema` / `ColumnSchema` | [schema-classes.md](./architecture/schema-classes.md) |
| Table and column class chains, `TableCommonRaw`, the Sheet container | [class-chains.md](./architecture/class-chains.md) |
