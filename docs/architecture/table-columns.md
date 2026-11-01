# Column inserts and column types

Map fragment. Sibling headings live in this folder.

A column is inserted only at a Table's end, in a band bounded to the Table and its head rows. Changing a column's type with `updateColumnType` has limits of its own.

## Column inserts are bounded to the Table

No queued operation inserts whole sheet columns. `TableRaw.appendColumn` is the one way to insert a column. It picks the index, the Table end past the inserts already queued, writes the new column's column ID, header and optional group heading, and returns the index; several on one Table land side by side in queue order. At gathering, `TableRaw.gatherInsertTableEndColumnsOperation` expands a Table's inserts into one `insertTableEndColumns`: an `insertRange` (`COLUMNS`) just past the Table's last column, from header −3 (the column ID row) down to the Table's last row, so the head rows move with the Table, then an `updateTable` with `fields: "range"` widening it, since an insert just past a Table leaves its range alone. Cells above and below that band stay put, and so does a Table below ([observed](../sheets-behaviour.md#column-inserts-bounded-to-a-table)). An insert starting past the grid's edge is refused ([observed](../sheets-behaviour.md#the-grids-edge)), so a Table ending at the grid's last column gets an `appendDimension` (`COLUMNS`) first, summed per sheet like growth's rows. Google heads the new column `Column <n>` until the queued header write renames it ([observed](../sheets-behaviour.md#column-inserts-bounded-to-a-table)), and its Table column type starts Automatic. The band's new cells take no format or validation live, and the fake leaves them blank too ([observed](../sheets-behaviour.md#column-inserts-bounded-to-a-table)). The flush grows the in-memory Table's end column by the insert count, matching the live Table. There is no mid-Table insert. How an insert moves and guards the Tables to its right: [stacked Tables](./stacked-tables.md#a-column-insert-pushes-the-tables-to-its-right).

## `updateColumnType` has three limits

One update per fetch: the flush clears the Table's fetched column list, so another update before a refetch throws. It refuses, before anything is sent, when the same flush inserts a column on that sheet, and on a Table with a validated column, because a restore resets its dropdown style and colours.
