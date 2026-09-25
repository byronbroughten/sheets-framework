# Architecture

Map fragments, one file per heading. Open the file the task needs.

| When | File |
| --- | --- |
| Chore homes, no registry, not tested | [chores.md](./architecture/chores.md) |
| `gatherRawRequest` | [raw-request-opening.md](./architecture/raw-request-opening.md) |
| Conditional format rules (order, prepend, identity, anchored formulas) | [conditional-format-rules.md](./architecture/conditional-format-rules.md) |
| Edit protections (edit warnings, edit locks, identity, stale refetch) | [edit-protections.md](./architecture/edit-protections.md) |
| Queuing a write, flush order, stale indexes, fills, formula writes, `findReplace`, discarding | [queued-writes.md](./architecture/queued-writes.md) |
| Active vs fetched, the working view, writes without a fetch, reads that need one | [working-view.md](./architecture/working-view.md) |
| A moved or extra Table | [table-placement.md](./architecture/table-placement.md) |
| Last data row, blank-row reuse | [blank-row.md](./architecture/blank-row.md) |
| Endpoint entry, `EndpointRun`, run report, selector behavior | [endpoint-dispatch.md](./architecture/endpoint-dispatch.md) |
| Sheets round trips | [round-trips.md](./architecture/round-trips.md) |
| Instantiation budget, the template-literal cliff | [type-check-cost.md](./architecture/type-check-cost.md) |
| Relative `<SN, CN>` vs `ColumnFullName` | [column-addressing.md](./architecture/column-addressing.md) |
| `SpreadsheetBaseSchema` / `SpreadsheetSchema` / `SheetSchema` / `ColumnSchema` | [schema-classes.md](./architecture/schema-classes.md) |
| Meta/primary class chains, `SheetCommonRaw` | [class-chains.md](./architecture/class-chains.md) |
