# Table placement: a moved or missing Table stops the run

Map fragment. Sibling headings live in this folder. The operator-facing word is **Table** in [`CONTEXT.md`](../../CONTEXT.md).

A run checks each managed Table against where the configs record it, and stops with a message to regenerate the configs when it isn't there. It never moves, rebuilds or picks a Table.

## The placement check

It runs per Table in `SpreadsheetRaw`'s post-fetch step, on each Let api access sheet whose placement strip rode the fetch:

- the sheet has a Table
- its header sits at the recorded row and column
- the column ID row (header −3) holds only blanks or this Table's own prefixed column IDs, and at least one ID, across the Table's columns as far as the fetch reached

Any failure stops the run and names the sheet. The same step stops on a Table met with only its header, once its header is in place ([blank row](./blank-row.md#a-table-met-with-only-its-header-stops-the-run)). For example, deleting a row above a Table or inserting a column to its left moves its header, and a Table moved out of the strip's sight reads as missing. A band of head rows shifted by an inserted row, with the header left in place, fails the column ID row test.

"Recorded" means the GID, `headerRowIndex` and `startColIndex` in the Table's `tableConfigs` entry. The edit trigger and the fetches sent before a Table's properties arrive aim there too, so a managed Table may sit lower on its sheet. Only a sheet the configs don't record falls back to the spot the framework creates Tables at (`TableOrigin.expected()`).

## The strip costs no round trip

`TableRaw.gatherFetchProperties` gathers one strip: from row 0 down to the recorded header, in the recorded start column. It overlaps the header cell, so Sheets returns the Table's properties with it, and it carries the column ID row's first cell. It rides the run's first fetch, so the check needs no fetch of its own; the old reclassifying sheet-properties fetch is gone ([round trips](./round-trips.md#table-ranges-and-table-bounded-reads)).

## More than one Table on a sheet

If a fetch finds more than one Table on a sheet with **Let api access**, it refuses and names those sheets, so you can delete the extras; it never guesses which one is managed. The exception is a sheet holding exactly one Table its configs record: the others beside it are unmanaged, and reaching the sheet's Table by GID reaches the recorded one. This refusal stays until several managed Tables may share a sheet (sheets-framework#89).

## Why the app never repairs a Table

It never moves or rebuilds a Table, because a Table that moved or multiplied usually means you restructured the sheet on purpose. No type can reach a range a person edits by hand, so this runtime check in `SpreadsheetRaw`'s post-fetch step is the only instrument there ([`docs/design/unrepresentable-disagreement.md`](../design/unrepresentable-disagreement.md), #9).

## Which sheets are checked

Everyday Table-placement and extra-Table checks use last-generate sheet GIDs, one regen behind the live box ([`docs/generated-data.md`](../generated-data.md#what-a-regeneration-runs-on-the-node-host)).
