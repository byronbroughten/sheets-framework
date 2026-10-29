# Cell writes, fills and formula writes

Map fragment. Sibling headings live in this folder.

Per-cell writes, column fills and formula writes: what each request carries, how a value, a formula and a colour merge, and which write wins when a fill and a cell write meet.

## Per-cell writes merge value, formula and colour

Most writes queue per cell and flush as one `updateCells` request each; a cell's queued entry holds an optional value, an optional formula, and an optional background colour (`RgbColor` in the Source tier, mapped to Google `Color` only inside `GoogleSheetsAPI`), merged across writes, so a value and a colour on one cell still cost one request and neither can cancel the other. A value and a formula are mutually exclusive on that same entry: writing one drops the other. Each request's field mask is built from what that entry actually holds, so a colour-only write can't blank the cell's value and a value-only write can't blank its colour — which is how a per-cell message queued after a column-wide fill keeps the fill's colour unless it names its own.

## Column fills

A column write instead queues one sheet-level _fill_ per contiguous run of rows: `updateAllCells` fills every table data row of a column in one `repeatCell`, and `updateWorkingCells` fills only the column's working rows, so stamping a selection costs no more requests than the selection is wide.

## Formula writes

`updateAllFormulas` / `updateWorkingFormulas` are the formula-column siblings of those fills, and `updateFormula` is the per-cell sibling of `updateValue`. Both are exposed from `ColumnRaw`/`CellRaw` up through Identified and Named under the same names. Named type-restricts the three formula methods to formula columns; Identified runtime-refuses a formula write on a non-formula column; `GoogleSheetsAPI` serializes a string that already starts with `=` as `pasteData` with `PASTE_FORMULA` (RFC4180-quoted fields, tab delimiter, identical text on every row of a fill) and does not consult `isFormula`. **`pasteData` is required, not a style choice: a formula sent as `updateCells` `formulaValue` is stored verbatim, but a Table reference in it (`SUM(T[Col])`, `SINGLE(T[Col])`) evaluates to a parse error. The same text pasted, or sent through the values API as `USER_ENTERED`, parses (measured live, and again in sheets-framework#53, 2026-10-01).** So re-writing a formula read back from `userEnteredValue` must also go through `pasteData`. Colour on the same change still serializes as today's `repeatCell` / `updateCells` colour path, so a formula-plus-colour flush is two requests in one `batchUpdate`. A formula string that does not start with `=` throws before anything is queued — the `=` is not prepended. A formula write does not mirror into local cell state, so a later read in the same run still sees the old effective value (colour-only writes are the precedent), and it does not invalidate reads.

A formula fill covers the Table body only, so it never reaches the head rows or a neighbour. `column.reference` (`household[ID]`) and `column.single` (`SINGLE(household[ID])`) build a Table reference from the live Table name and header cell, so formula text stays right after a rename without regenerating; there is no formula tagged template. `column.profile.topFormula` is the top data row's formula text, the "before" a formula-rewrite dry run shows.

## The last queued write wins between fills and cell writes

**The write queued last wins, whether it is a fill or a cell write, field by field.** Column fills still go out before cell writes, so a cell write queued after a fill sits on top of it; a fill queued after a cell write wins because queuing it erases the fields it sets from every cell entry already queued in its row range on that column. A value or formula clears the entry's value and formula, one slot; a colour clears its colour; an entry left with no fields is dropped. The order costs no extra request. A fill's last row is snapshotted when queued, so it never reaches a row appended after it. State is still mirrored row by row for _value_ fills — only the request collapses. `updateAllCells` and `updateAllFormulas` throw on a sheet that `TableRaw.removeRowsExcept` has pruned to a selection, since a whole-column fill ignores working rows and would reach exactly the rows the prune excluded.
