# Grow a Table by insert plus widen

A Table grows by inserting rows limited to its own columns and widening its range, never by `appendCells`, which inserts whole sheet rows, splits loose cells and fails on a header-only Table; using `appendCells` where no neighbour shares the rows was rejected as two growth paths with two sets of failure modes. Each Table's growth is one queued `appendTableRows`, merged by `tableId` and expanded at gathering into one `appendDimension` per sheet, the insert, the widen and the format copy. The stale flag this decision also set on a grown Table and the Tables below it was removed, since every row index counts from its Table and stays valid after growth (byronbroughten/sheets-framework#120).

Source: byronbroughten/sheets-framework#64
