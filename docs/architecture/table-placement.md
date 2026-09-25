# Table placement: a moved or extra Table stops the run

Map fragment. Sibling headings live in this folder. The operator-facing word is **Table** in [`CONTEXT.md`](../../CONTEXT.md).

A run checks where each Let api access sheet's Table starts, and refuses to go on when it has drifted or has company. It never moves, rebuilds or picks a Table.

## A moved Table

Deleting a row above a Table or inserting a column to its left moves it, so the app checks where it starts on every run and refuses to go on if it has drifted, naming where the Table is and where it belongs.

## More than one Table on a sheet

If a fetch finds more than one Table on a sheet with **Let api access**, it refuses the same way and names those sheets, so you can delete the extras; it never picks one for you.

## Why the app never repairs a Table

It never moves or rebuilds a Table, because a Table that moved or multiplied usually means you restructured the sheet on purpose. No type can reach a range a person edits by hand, so this runtime check in `SpreadsheetRaw`'s post-fetch step is the only instrument there ([`docs/design/unrepresentable-disagreement.md`](../design/unrepresentable-disagreement.md), #9).

## Which sheets are checked

Everyday Table-placement and extra-Table checks use last-generate sheet GIDs, one regen behind the live box ([`docs/generated-data.md`](../generated-data.md#what-a-regeneration-runs-on-the-node-host)).
