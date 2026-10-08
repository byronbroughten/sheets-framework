# Table placement: a moved or missing Table stops the run

Map fragment. Sibling headings live in this folder. The operator-facing word is **Table** in [`CONTEXT.md`](../../CONTEXT.md).

A run checks each managed Table against where the configs record it, and stops with a message to regenerate the configs when it isn't there. It never moves, rebuilds or picks a Table.

## The placement check

It runs per recorded Table in `SpreadsheetRaw`'s post-fetch step, on each Let api access Table whose placement strip rode the fetch:

- the Table is on its sheet: found by its recorded `tableId`, or, on a sheet the configs record only it on, as that sheet's one Table
- its header sits at the recorded row and column
- the column ID row (header −3) holds only blanks or this Table's own prefixed column IDs, and at least one ID, across the Table's columns as far as the fetch reached

Any failure stops the run and names the sheet; the recorded start position tells apart Tables that share one. The same step stops on a Table met with only its header, once its header is in place ([blank row](./blank-row.md#a-table-met-with-only-its-header-stops-the-run)). For example, deleting a row above a Table or inserting a column to its left moves its header, and a Table moved out of the strip's sight reads as missing. A band of head rows shifted by an inserted row, with the header left in place, fails the column ID row test.

"Recorded" means the GID, `headerRowIndex` and `startColIndex` in the Table's `tableConfigs` entry. The edit trigger and the fetches sent before a Table's properties arrive aim there too, so a managed Table may sit lower on its sheet. Only a sheet the configs don't record falls back to the spot the framework creates Tables at (`TableOrigin.expected()`).

## Checked once per run

A Table's position is checked on the fetch that first tells the run where it is. After that the run moves it itself: growth pushes the Tables below whose columns it overlaps down, a column insert pushes the Tables to its right along, and a row delete pulls the Tables below up ([queued writes](./queued-writes.md)). So a later fetch in the same run doesn't take a pushed neighbour it had already fetched as moved. A neighbour it hadn't fetched yet isn't shifted in state, so it is still judged against its recorded origin.

**Accepted gap:** the next run, or a same-run first fetch of a neighbour already pushed, stops on it. A Table one of these pushed sits off its recorded origin, and the run stops on it as moved or missing until the configs are regenerated.

## The strip costs no round trip

`TableRaw.gatherFetchProperties` gathers one strip: from row 0 down to the recorded header, in the recorded start column. It overlaps the header cell, so Sheets returns the Table's properties with it, and it carries the column ID row's first cell. It rides the run's first fetch, so the check needs no fetch of its own; the old reclassifying sheet-properties fetch is gone ([round trips](./round-trips.md#table-ranges-and-table-bounded-reads)).

## Head rows that sit on another Table stop the config sync

Google refuses one Table's range over another's, but it can't see head rows, so it accepts a managed Table whose head rows sit on another Table. Head-row writes would then land in the other Table's cells. Nothing the framework does creates that overlap: growth, row deletes and the column insert move a neighbour together with its head rows. Only the operator's arrangement can, and every change to a managed Table's layout passes through a config sync.

So a config sync checks it, in `ConfigCoordinator._syncConfigSheetRows`, once the Table Config rows are read and before anything is flushed or a config file is returned. Each Table ticked **Let api access** on the live sheet, this sync's ticks included, has its head rows compared with the range of every other Table on its sheet, managed or not. The comparison uses geometry alone and reads no cells. A managed Table's head rows on another managed Table's head rows also sit on that Table's range, since a range starts at its header row, so one message covers both. Every overlap goes into one throw, naming the Table, its head rows' 1-based rows and the Table under them.

The head rows' extent runs from `tableLayout`'s largest offset down to just above the header (`headRows.topIndex`, `TableOrigin.topHeadSheetRowIndex`). The column insert's shift and split check uses the same extent.

**Accepted gap:** an unmanaged Table added on top of a managed Table's head rows, with no regeneration since, isn't caught.

## Several Tables on a sheet

Several managed Tables may share a sheet, side by side or stacked (sheets-framework#89). Each is reached by its recorded `tableId`, before its properties arrive too, and gathers its own strip at its own recorded origin. A GID reaches a Table only on a sheet the configs record one Table on: the sheet's only Table, else the recorded one, so unmanaged Tables may sit beside it.

A grid fetch returns a sheet's `tables` only for the ranges it overlaps, so a Table it leaves out keeps its state. Only a fetch covering a whole sheet, the sheet-properties fetch or `fetchSheetUsedGrid`, removes a Table absent from it.

## Why the app never repairs a Table

It never moves or rebuilds a Table, because a Table that moved or multiplied usually means you restructured the sheet on purpose. No type can reach a range a person edits by hand, so this runtime check in `SpreadsheetRaw`'s post-fetch step is the only instrument there ([`docs/design/unrepresentable-disagreement.md`](../design/unrepresentable-disagreement.md), #9).

## Which sheets are checked

Everyday Table-placement checks use last-generate sheet GIDs, one regen behind the live box ([`docs/generated-data.md`](../generated-data.md#what-a-regeneration-runs-on-the-node-host)).
