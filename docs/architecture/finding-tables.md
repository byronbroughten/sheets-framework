# Finding a managed Table wherever it sits

Map fragment. Sibling headings live in this folder.

Which read still reaches a managed Table after it moves, and what each costs, probed live on the dev spreadsheet on 2026-10-08 (sheets-framework#130); the observations are in the Sheets index's [Finding Tables](../sheets-behaviour.md#finding-tables). The header zone is the top rows of a sheet, across every column, which a run reads to find its Tables (sheets-framework#129). Recommendation: keep it, and deepen it when stacked Tables return.

## The Probe130 tab, its moves and its routes

A throwaway `Probe130` tab held two Tables: `layoutAbove` with its header on row 4 (B:C), and `p130Target` stacked below it with its column ID row on row 10 and its header on row 13 (B:D). The tab also held a named range over `p130Target`'s head rows and header (B10:D13), and developer metadata on its header row and on its first column. Four moves ran in turn, each on the layout the last one left, and every route was tried after each:

- **Append on the Table above**: a row inserted inside `layoutAbove`, which grows it and pushes `p130Target` down.
- **Column insert to its left**: column A.
- **Row delete above it**: the first row below `layoutAbove`.
- **Cut and paste**: the API's `cutPaste` of `p130Target`'s head rows and range to H41. This stands in for an operator's cut and paste.

A route reaches the Table when `p130Target` comes back in the response's `tables`. The routes:

0. **The placement strip** the first fetch gathered before sheets-framework#129: column B from row 1 to the recorded header row, plus the column ID row, in one request.
1. The Table's name or a Table reference as the `a1Range`.
2. The named range over the head rows.
3. Developer metadata, on the header row or the first column.
4. The whole spreadsheet's `sheets.tables`.
5. The header zone at 4, 20 and 100 rows, each starting at row 1.

Costs were measured before the first move, from a laptop through the Node host's credential, so the latencies are not from Apps Script; the pacing is in the [timing line](../sheets-behaviour.md#reads-that-find-a-table), under Google's [read quota](../sheets-behaviour.md#a-round-trips-cost). The tab was deleted afterwards.

## What reaches the Table after each move

- **A Table reference is not `a1Range` syntax**, and one unparseable `a1Range` fails the whole request ([observed](../sheets-behaviour.md#naming-a-table-in-a-data-filter)), so a deleted named range would fail every fetch it rides.
- **Only the whole-spreadsheet fetch, the named range and a 100-row zone reached the Table after every move** ([zone and strip](../sheets-behaviour.md#which-tables-a-grid-fetch-returns), [named range](../sheets-behaviour.md#developer-metadata-and-named-ranges-under-moves)). The placement strip lost it on the first move, and the 4-row zone never saw it, its header sitting on row 13: the stacked case sheets-framework#129 leaves for later.
- **Developer metadata lost the Table only on the cut and paste**, and its lookups return whole rows or used columns, the Tables above included ([observed](../sheets-behaviour.md#developer-metadata-and-named-ranges-under-moves)).
- **The named range followed every move because the cut covered all of it.** Two cases are untested: a cut of the Table's range that leaves its head rows behind, and an operator deleting the named range.

## The premise under the header zone holds

A grid fetch returned only the Tables its range overlapped, for every range tried ([observed](../sheets-behaviour.md#which-tables-a-grid-fetch-returns)), which is what `round-trips.md` relies on. A range over a Table's head rows alone returned no Tables, so a read must reach the header row.

## What each route costs

The latencies and response sizes are in the [timing line](../sheets-behaviour.md#reads-that-find-a-table).

- **Every route that rides a grid fetch costs about what the placement strip does.** The differences are within one round trip's noise, and only removing a round trip moves the number ([round trips](./round-trips.md)). Developer metadata ran about 50ms slower.
- **The zone's size follows the used cells inside it, not its depth.** 100 rows cost what 20 did until the paste put used rows in range. On a live sheet with body rows inside the zone, a deep zone pulls those rows across every used column.
- **Route 4 carries no cells, but it is a round trip of its own**, and its size grows with the spreadsheet's sheets and Tables. A per-sheet version, `getByDataFilter` with the whole sheet as the `a1Range` and no `data` in the mask, is smaller, but still costs its own round trip, since adding `data` back to the mask would return the whole sheet's cells.

## Keep the header zone, and deepen it for stacked Tables

**A zone deep enough for the lowest stacked header is the cheapest route that reaches stacked Tables.** A deeper zone reaches the stacked header on the same round trip. The added cost is the used cells it covers.

- **The whole-spreadsheet fetch finds a Table anywhere, but it adds a round trip to every run.** The edit dispatch would go from two round trips to three, at [a round trip's cost](../sheets-behaviour.md#a-round-trips-cost).
- **A named range per managed Table, over its head rows and header, is the route for finding a Table anywhere without an extra round trip.** It is the only route that rides the fetch and survives every move without depending on how far down the Table sits. It costs `gen:configs` creating and maintaining one named range per Table. A deleted named range would fail every fetch it rides, so it would need a fallback first, and the head-rows-left-behind cut needs a test.
- **Developer metadata loses the most.** It loses the Table on a cut and paste, it reads whole rows or columns, and it is no cheaper.
