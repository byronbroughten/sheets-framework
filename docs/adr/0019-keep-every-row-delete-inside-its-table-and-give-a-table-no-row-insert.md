# Keep every row delete inside its Table, and give a Table no row insert

Deleting a sheet row cuts any other Table on the same rows, so every row delete is a `deleteRange` that shifts rows up across the Table's own columns, sparing a neighbouring Table and any cells beside it, and no sheet-row delete survives, not even on a one-Table sheet. A Table gets no row insert either: it grows only at its end, and an insert, if one is ever wanted, would be an `insertRange` limited to the Table's columns. Both follow the standing stance that no queued operation touches whole sheet rows or columns; the growth this decision kept on `appendCells` later moved to insert plus widen for the same reason ([ADR 0013](./0013-grow-a-table-by-insert-plus-widen.md)).

Source: byronbroughten/sheets-framework#52
