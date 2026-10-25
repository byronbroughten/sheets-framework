# Grow a Table by insert plus widen, and flag the layout it moves as stale

A Table grows by inserting rows limited to its own columns and widening its range, never by `appendCells`, which inserts whole sheet rows, splits loose cells and fails on a header-only Table. Using `appendCells` when no neighbour shares the sheet's rows was rejected, because it would mean two growth paths with two sets of failure modes and tests. Each Table's growth is one queued `appendTableRows`, merged by `tableId` and expanded at gathering into one `appendDimension` per sheet, the insert, the widen and the format copy, with values and formulas left to the existing fills. After a flush that grows a Table, that Table and every overlapping Table below it are flagged stale, and a later row-index conversion or queued write on one throws, saying a refetch is needed, since a run flushes once and nothing restores the layout.

Source: byronbroughten/sheets-framework#64
