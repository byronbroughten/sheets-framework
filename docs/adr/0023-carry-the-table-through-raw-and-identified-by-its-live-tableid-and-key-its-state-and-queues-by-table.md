# Carry the Table through Raw and Identified by its live `tableId`, and key its state and queues by Table

Raw addresses a Table by its live `tableId`, as it does a sheet by GID, because bounds go stale after growth and aren't identity; rows and columns count from the Table, and sheet coordinates are built only at gathering, against live bounds. Sheet state keeps the per-sheet facts (title, `rowCount`, conditional formats, edit protections), and a `tableId`-keyed Table state holds bounds, cells and the Table's queues, so Tables sharing a sheet never collide. An operation carries a `tableId` only where its request takes one (`appendCells`, `updateTable`); every other one is bounded by whoever builds it from the Table-keyed queue. The sheet gets no Identified class, since its every member would only pass through to Raw.

Source: byronbroughten/sheets-framework#54, byronbroughten/sheets-framework#62
