# Finding a managed Table wherever it sits

Map fragment. Sibling headings live in this folder.

Which read still reaches a managed Table after it moves, and what each costs, measured live on the dev spreadsheet on 2026-10-08 (sheets-framework#130). The header zone is the top rows of a sheet, across every column, which a run reads to find its Tables (sheets-framework#129). Recommendation: keep it, and deepen it when stacked Tables return.

## The Probe130 tab, its moves and its routes

A throwaway `Probe130` tab held two Tables: `layoutAbove` with its header on row 4 (B:C), and `p130Target` stacked below it with its column ID row on row 10 and its header on row 13 (B:D). The tab also held a named range over `p130Target`'s head rows and header (B10:D13), and developer metadata on its header row and on its first column. Four moves ran in turn, each on the layout the last one left, and every route was tried after each:

- **Append on the Table above**: a row inserted inside `layoutAbove`, which grows it and pushes `p130Target` down.
- **Column insert to its left**: column A.
- **Row delete above it**: the first row below `layoutAbove`.
- **Cut and paste**: the API's `cutPaste` of `p130Target`'s head rows and range to H41. This stands in for an operator's cut and paste.

A route reaches the Table when `p130Target` comes back in the response's `tables`. Route 0 is the placement strip as the first fetch gathers it today: column B from row 1 to the recorded header row, plus the column ID row, in one request. Every header zone starts at row 1. Costs were measured before the first move.

A one-off Node script made the calls with the Node host's credential, so its latencies are from a laptop and not from Apps Script. Each route was timed 12 times, paced under the 60-reads-a-minute quota. The tab was deleted afterwards.

## What reaches the Table after each move

| Route | Before | Append above | Column insert left | Row delete above | Cut and paste |
| --- | --- | --- | --- | --- | --- |
| 0. Placement strip, at the recorded origin | yes | no | no | no | no |
| 1. Table name or reference as `a1Range` | 400 | 400 | 400 | 400 | 400 |
| 2. Named range over the head rows | yes | yes | yes | yes | yes |
| 3. Developer metadata, header row or first column | yes | yes | yes | yes | no |
| 4. Whole-spreadsheet `sheets.tables` | yes | yes | yes | yes | yes |
| 5a. Header zone, 4 rows | no | no | no | no | no |
| 5b. Header zone, 20 rows | yes | yes | yes | yes | no |
| 5c. Header zone, 100 rows | yes | yes | yes | yes | yes |

- **A Table reference is not `a1Range` syntax.** `p130Target`, `p130Target[ID]` and `Probe130!p130Target` each fail with 400 "Unable to parse range".
- **One unparseable `a1Range` fails the whole request**, not only its own filter. A two-filter `getByDataFilter` with one valid range and one unknown name returned 400, so a deleted named range would fail every fetch it rides.
- **Developer metadata follows row and column inserts and deletes**, against the issue's expectation, because it rides the row or column itself. It loses the Table only when cells move without their rows, as in a cut and paste. The header-row lookup returns the whole sheet row. The column lookup returns the whole used column, so it brings back the Tables above as well.
- **The named range followed every move, cut and paste included**, because the cut covered all of it. Two cases are untested: a cut of the Table's range that leaves its head rows behind, and an operator deleting the named range.
- **The 4-row zone never saw `p130Target`**, since its header sat on row 13. That is the stacked case sheets-framework#129 leaves for later.

## The premise under the header zone holds

`round-trips.md`'s claim that a grid fetch returns only the Tables it overlaps held for every range tried:

- A range over `layoutAbove`'s header cell alone returned `layoutAbove` only.
- A range that overlaps no Table returned no Tables.
- A range over `p130Target`'s three head rows, ending just above its header, returned no Tables. Head rows are not part of a Table, so a read must reach the header row.
- The whole sheet as one range returned every Table on it.

## What each route costs

| Route | Median ms | Min–max ms | Response bytes |
| --- | --- | --- | --- |
| 0. Placement strip | 201 | 191–228 | 5,257 |
| 2. Named range | 210 | 194–241 | 2,742 |
| 3. Developer metadata, header row | 250 | 225–940 | 2,741 |
| 3. Developer metadata, first column | 254 | 238–1,864 | 3,953 |
| 4. Whole-spreadsheet `sheets.tables` | 188 | 156–233 | 13,392 |
| 5a. Header zone, 4 rows | 202 | 186–231 | 2,567 |
| 5b. Header zone, 20 rows | 202 | 187–230 | 6,672 |
| 5c. Header zone, 100 rows | 214 | 188–250 | 6,672 |

- **Every route that rides a grid fetch costs about what the placement strip does.** The differences are within one round trip's noise, and only removing a round trip moves the number ([round trips](./round-trips.md)). Developer metadata ran about 50ms slower.
- **The zone's size follows the used cells inside it, not its depth.** Rows past the last used row add nothing: 100 rows cost what 20 did until the paste put rows 41–47 in range (8,270 bytes). On a live sheet with body rows inside the zone, a deep zone pulls those rows across every used column.
- **Route 4 carries no cells, but it is a round trip of its own.** Its size grows with the spreadsheet's sheets and Tables: 13 sheets came to 13 KB. A per-sheet version, `getByDataFilter` with the whole sheet as the `a1Range` and no `data` in the mask, returned the Layout sheet's three Tables in under 2 KB. That still costs its own round trip, since adding `data` back to the mask would return the whole sheet's cells.

## Keep the header zone, and deepen it for stacked Tables

**A zone deep enough for the lowest stacked header is the cheapest route that reaches stacked Tables.** A deeper zone reaches the stacked header on the same round trip. The added cost is the used cells it covers.

- **The whole-spreadsheet fetch finds a Table anywhere, but it adds a round trip to every run.** The edit dispatch would go from two round trips to three, at [a round trip's cost](./round-trips.md#the-chokepoints).
- **A named range per managed Table, over its head rows and header, is the route for finding a Table anywhere without an extra round trip.** It is the only route that rides the fetch and survives every move without depending on how far down the Table sits. It costs `gen:configs` creating and maintaining one named range per Table. A deleted named range would fail every fetch it rides, so it would need a fallback first, and the head-rows-left-behind cut needs a test.
- **Developer metadata loses the most.** It loses the Table on a cut and paste, it reads whole rows or columns, and it is no cheaper.
