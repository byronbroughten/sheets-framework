# Work in the concept's own frame; convert to the platform's only at the boundary

Design reasoning. The one-line principle lives in [`docs/design.md`](../design.md); this file holds the argument and its instances.

## The argument

Code reasons about the thing it's about, so its indexes should count from that thing, not from wherever the platform happens to put it. Convert to the platform's frame once, where requests leave, and give the platform's frame a branded type, so a number from one frame can't be passed where the other is expected. A frame converted at every call site is a frame that's wrong at one of them.

## Instances

[ADR 0030](../adr/0030-count-rowindex-from-the-tables-first-body-row-and-brand-the-sheet-row.md) counts `rowIndex` from the Table's first body row and converts it only where requests leave, into a branded `SheetRowIndex` (a column, a `SheetColIndex`), so a Table-relative index can't reach Google unconverted. [ADR 0027](../adr/0027-give-each-table-its-own-head-rows-at-fixed-offsets-above-its-header.md) places the head rows at offsets from the Table's own header rather than at sheet positions, so a Table can sit anywhere its head fits. [ADR 0023](../adr/0023-carry-the-table-through-raw-and-identified-by-its-live-tableid-and-key-its-state-and-queues-by-table.md) keeps Raw's state in the Table's frame, so two Tables on one sheet never collide in state and a Table moved down its sheet changes no index. Growth is the instance that tests the boundary: a Table pushed down by a growth above it keeps every `rowIndex`, because each later operation in the batch converts against the layout the earlier ones leave, and only the conversion moves ([ADR 0013](../adr/0013-grow-a-table-by-insert-plus-widen.md) drops growth's stale flag because every index counts from its Table).
