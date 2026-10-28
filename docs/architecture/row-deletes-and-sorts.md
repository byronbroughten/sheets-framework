# Row deletes and sorts stay within the Table

Map fragment. Sibling headings live in this folder.

A row delete and a sort each act on one Table's columns at its live bounds, never on whole sheet rows, so a side-by-side neighbour stays put.

## Row deletes shift only the Table's columns

No queued operation deletes whole sheet rows. `RowRaw.delete` sets a delete flag in its Table's write queue, keyed by row index, so a delete on one Table never flags a neighbour's row. The flush gathers each flagged row with `RowCommonRaw.gatherDeleteTableRowsOperation` into a `deleteTableRows` whose range is that row across the Table's columns at its live bounds, and the adapter sends it as a `deleteRange` shifting rows up (sheets-framework#52). Table writes are gathered before row writes, so a Table-end column insert queued in the same flush is already counted in the band: a band one column short would cover part of the Table, which Google refuses. A side-by-side Table and loose cells beside the Table stay where they are; a Table below in the same columns moves up with the shift, and a sort on it in the same batch follows it ([stacked Tables](./stacked-tables.md#a-row-delete-pulls-the-tables-below-up)).

## Sorts are bounded to the Table body

`TableRaw.queueSort` queues one sort per Table, gathered into a `sortTable` whose range is the body rows the same flush's deletes leave, across the Table's columns at its live bounds, and sent as a `sortRange` after those deletes (sheets-framework#95). A Table pulled up by deletes above it sorts at its rows after them. The head rows, a side-by-side Table and the rows below stay put, and the flush marks the sorted Table's row indexes stale.
