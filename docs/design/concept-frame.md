# Work in the concept's own frame; convert to the platform's only at the boundary

Design reasoning. The one-line principle lives in [`docs/design.md`](../design.md); this file holds the argument and its instances.

## The argument

Code reasons about the thing it's about, so its indexes should count from that thing, not from wherever the platform happens to put it. Convert to the platform's frame once, where requests leave, and give the platform's frame a branded type, so a number from one frame can't be passed where the other is expected. A frame converted at every call site is a frame that's wrong at one of them.

## Instances

`rowIndex` counts from the Table's first body row, through Named, Raw and the write queue, rather than from the sheet's top. It is converted to a sheet row only at gathering, at the edit trigger's entry and in operator messages, and a sheet row is a branded `SheetRowIndex` (a column, a `SheetColIndex`), so a Table-relative index can't reach Google unconverted (sheets-framework#55). Raw holds every row and column state in Table-relative indexes and builds sheet coordinates only when the flush gathers a request, against the Table's live bounds, so two Tables on one sheet never collide in state and a Table moved down its sheet changes no index (sheets-framework#62). Growth is the instance that tests the boundary: a Table pushed down by a growth above it keeps every `rowIndex`, because each later operation in the batch converts against the layout the earlier ones leave, and only the conversion moves (sheets-framework#64).
