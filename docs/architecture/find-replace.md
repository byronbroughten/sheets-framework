# `findReplace`

Map fragment. Sibling headings live in this folder.

The one write that names no coordinates: its fields, its scope, its slot in the flush, and the reads it invalidates.

## `findReplace` invalidates every read

**`findReplace` is the one write that names no coordinates, and it pays for that by invalidating every read.** It carries Google's own fields — `find`, `replacement`, `matchCase`, `matchEntireCell`, `searchByRegex`, `includeFormulas` — and a scope of a range or the whole spreadsheet. A Table's or a column's is queued on the Table and bounded at gathering to the body rows at the Table's live bounds, so it never reaches the head rows or another Table, and `allSheets` stays a chore tool; there is no sheet scope (sheets-framework#95). Both strings stay plain: the value config governs neither, because Google matches the cell's text. It rides the same queue and the same dry run, and sits **after** cell writes because it reads the text as it stands mid-batch, so text queued beside it must land first. Two consequences follow from matching by content. Nothing can be mirrored into local state, since what matched is unknowable locally, so the flush marks **every** sheet's fetched cells stale — one rule rather than a scope-dependent one — and a read afterwards throws naming the findReplace until the next fetch clears it. `includeFormulas` is the text-level rewrite of a formula already in place; stamping a new formula by coordinate is a [formula write](./cell-writes.md#formula-writes), which does not invalidate reads. `validateDataNotFormula` still refuses to _write a value_ onto a formula column, because a formula written as a value misbehaves in Apps Script.
