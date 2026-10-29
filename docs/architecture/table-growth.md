# Growing a Table

Map fragment. Sibling headings live in this folder.

Appends grow a Table by one `appendTableRows` per Table: what it sends, which row it models the new rows on, how a lone blank row is reused, and what it leaves beside the Table.

## Growth is one `appendTableRows` per Table

`RowRaw.append` flags the row in its Table's write queue, so appends are grouped by `tableId` before any operation is built, and two Tables on one sheet grow apart. At gathering, `TableRaw.gatherAppendTableRowsOperation` expands a Table's appends into one `appendTableRows` (sheets-framework#59, #61):

1. `appendDimension`: one per sheet, sent first, sized `max(0, needed bottom − rowCount)` summed over every growth on the sheet. The sum is held on the sheet's write queue, and the local `rowCount` rises with it, so two growths never both count the same rows. An insert can't reach past the grid's last row, which is why this goes first.
2. Per Table, bottom Table first within a sheet: an `insertRange` (`ROWS`) over the Table's columns, below the last body row; an `updateTable` with `fields: "range"` to widen it, since an insert below a Table leaves its range alone; a `copyPaste` `PASTE_FORMAT` from the model row onto the new rows, which carries number formats and colours and extends a conditional-format rule ending at the last row; and a `copyPaste` `PASTE_DATA_VALIDATION` over each run of untyped columns only, since Google refuses cell validation in a typed column.
3. Then the value fills (`updateCells` / `repeatCell`) and formula fills (`pasteData`) already queued on the new rows.

The model row is the body row just above the first queued append. It is taken from the queued rows, not from the Table's end, because a same-run re-fetch resets the end while keeping the queued appends. The column types are the fetched ones, with any queued in the same flush laid over them, since `setTableColumnProperties` is sent first; a Table column with no fetched type refuses the growth and asks for a refetch. A row appended and deleted in the same flush still grows the Table and then goes out with the row deletes, so the appended rows after it keep their indexes. Growth is gathered before the Table-end column inserts are counted, because it is sent before them. `rowCount` and the column types ride the standing fetch masks, so growth adds no round trip.

## Growing from a lone blank row

**Growing a Table whose only body row is blank writes that row in place and inserts the rest.** Growth by n then sends the first row's fills into the blank row and an `appendTableRows` for the other n − 1, modelled on the blank row; growth by one sends only the fills. The reuse is decided at append time, not at gathering: `TableIdentified.appendRowDefault` hands out the blank row instead of calling `RowRaw.append`, so Raw only ever sees the n − 1 real appends, and their model row (the one above the first of them) is the blank row. Raw can't decide it, because the blank test skips formula and feedback columns, which only Identified resolves. When reuse applies and when it is refused: [blank row](./blank-row.md#the-append-path-reuses-the-blank-row).

## Growth touches only the Table's columns

Growth touches only the Table's own columns: a side-by-side neighbour and the cells between get no blank rows. Loose cells below the Table in its columns move down with the new rows, while cells beside them stay, so a loose block sticking out past the Table's columns is split. That split is accepted: loose cells aren't a Table, so nothing checks for them (sheets-framework#57). A Table below in the same columns moves down: [stacked Tables](./stacked-tables.md#growth-pushes-the-tables-below-down).
