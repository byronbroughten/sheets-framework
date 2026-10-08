# Table placement: a missing Table, or one outside the header zone, stops the run

Map fragment. Sibling headings live in this folder. The operator-facing words are **Table** and **Header zone** in [`CONTEXT.md`](../../CONTEXT.md).

Every managed Table keeps its header row in the sheet's header zone: its top rows, across every column, `tableLayout.headerZoneDepth` deep (4 by default: the column ID row, two group-heading rows, the header row). A Table moves freely within the zone, so the configs record no position, only its GID and `tableId`. A run that meets a managed Table outside the zone, or misses one, stops and names it. It never moves, rebuilds or picks a Table.

## The placement check

It runs per recorded Table in `SpreadsheetRaw`'s post-fetch step (`SpreadsheetTableValidatorRaw.tablePlacements`):

- the Table is on its sheet: found by its recorded `tableId`, or, on a sheet the configs record only it on, as that sheet's one Table. Judged only on a fetch the sheet's zone rode, since only the zone is sure to bring it.
- its header row is in the zone, with room for its head rows above it (`headerZone.holdsHeaderRow`). With the default depth that means row 4. Judged on every fetch that brings the Table, since it reads geometry alone.
- the column ID row (header −3) holds only blanks or this Table's own prefixed column IDs, and at least one ID, across the Table's columns. Judged on a fetch the zone rode, since the zone carries that row.

Any failure stops the run, naming the sheet and the Table: "must have its header row on row 4", or "has no Table … with its header row on row 4". The same step stops on a Table met with only its header, once its header is in the zone ([blank row](./blank-row.md#a-table-met-with-only-its-header-stops-the-run)). For example, inserting a row above a Table pushes its header out of the zone, and the zone fetch no longer sees it, so the run reads it as missing. A band of head rows shifted by an inserted row, with the header left in place, fails the column ID row test.

## Moves within the zone need no regeneration

A column insert or delete to a Table's left, or a move along its rows, leaves its header in the zone, and the next run finds it there. The run's own column insert pushes the Tables to its right along, within the zone, so a later run reads them where they landed ([queued writes](./queued-writes.md)). No check is once-per-run: a Table moved within the zone is fine on every fetch.

Before a Table's properties arrive, the few reads and writes that need a position aim at `TableOrigin.expected()`, the spot the framework creates Tables at. The column ID row's fetch, the one read every first fetch needs, rides the zone instead.

## The zone costs no round trip

`TableRaw.gatherFetchProperties` gathers the zone once per sheet: rows 0 to the zone's depth, every column. Sheets returns every Table a grid range overlaps ([confirmed live](./finding-tables.md#the-premise-under-the-header-zone-holds)), so every Table whose header sits in the zone arrives with it, with its head rows. It rides the run's first fetch, so the check needs no fetch of its own.

The zone costs the used cells it covers, not its depth ([finding Tables](./finding-tables.md#what-each-route-costs)). On the dev Layout tab it came back smaller than the strips it replaced: 1,686 bytes for both Tables against 1,882 for two placement strips and their column ID rows.

## Generation refuses a ticked Table outside the zone

A config sync checks each Table ticked **Let api access**, this sync's ticks included, in `ConfigCoordinator._syncConfigSheetRows`, right after the head-row overlap check below and before anything is flushed or a config file is returned (`SpreadsheetTableValidatorRaw.validateHeadersInZone`). Every Table outside the zone goes into one throw, naming it. So the configs never record a Table that every run would stop on.

## Head rows that sit on another Table stop the config sync

Google refuses one Table's range over another's, but it can't see head rows, so it accepts a managed Table whose head rows sit on another Table. Head-row writes would then land in the other Table's cells. Nothing the framework does creates that overlap: growth, row deletes and the column insert move a neighbour together with its head rows. Only the operator's arrangement can, and every change to a managed Table's layout passes through a config sync.

So a config sync checks it, in `ConfigCoordinator._syncConfigSheetRows`, once the Table Config rows are read and before anything is flushed or a config file is returned. Each Table ticked **Let api access** on the live sheet, this sync's ticks included, has its head rows compared with the range of every other Table on its sheet, managed or not. The comparison uses geometry alone and reads no cells. A managed Table's head rows on another managed Table's head rows also sit on that Table's range, since a range starts at its header row, so one message covers both. Every overlap goes into one throw, naming the Table, its head rows' 1-based rows and the Table under them.

The head rows' extent runs from `tableLayout`'s largest offset down to just above the header (`headRows.topIndex`, `TableOrigin.topHeadSheetRowIndex`). The column insert's shift and split check uses the same extent.

**Accepted gap:** an unmanaged Table added on top of a managed Table's head rows, with no regeneration since, isn't caught.

## Several Tables on a sheet

Several managed Tables may share a sheet side by side, each with its header in the zone (sheets-framework#89). Each is reached by its recorded `tableId`, before its properties arrive too, and one zone fetch brings them all. A GID reaches a Table only on a sheet the configs record one Table on: the sheet's only Table, else the recorded one, so unmanaged Tables may sit beside it.

**Stacked managed Tables wait for a deepenable zone.** A Table below another has its header below the zone, so generation refuses it. Growth, row deletes and the column insert still push and pull the Tables below, and the fake covers it with an unmanaged Table under the managed ones. When stacked Tables return, deepen the zone to just below the lowest stacked header in the layout, not a blanket large number ([finding Tables](./finding-tables.md#keep-the-header-zone-and-deepen-it-for-stacked-tables)).

A grid fetch returns a sheet's `tables` only for the ranges it overlaps, so a Table it leaves out keeps its state. Only a fetch covering a whole sheet, the sheet-properties fetch or `fetchSheetUsedGrid`, removes a Table absent from it.

## The edit trigger reads the live Table

`Api.isSuspectedApiCall` costs no fetch: a TRUE or FALSE on a recorded sheet, on the row an action row in the zone can sit on (`SpreadsheetSchema.mayHoldActionCell`). `Api.handleSheetEdit` then calls `SpreadsheetIdentified.fetchTableWithActionCell`, which gathers each recorded Table's prerequisites on that sheet, all riding one zone fetch, and picks the Table whose live action row and columns hold the edited cell (`TableRaw.holdsActionCellAt`). The dispatch stays at one read and one write ([round trips](./round-trips.md)).

## Why the app never repairs a Table

It never moves or rebuilds a Table, because a Table that moved out of the zone or multiplied usually means you restructured the sheet on purpose. No type can reach a range a person edits by hand, so this runtime check in `SpreadsheetRaw`'s post-fetch step is the only instrument there ([`docs/design/unrepresentable-disagreement.md`](../design/unrepresentable-disagreement.md), #9).

## Which sheets are checked

Everyday Table-placement checks use last-generate sheet GIDs, one regen behind the live box ([`docs/generated-data.md`](../generated-data.md#what-a-regeneration-runs-on-the-node-host)).
