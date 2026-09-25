# The payload is not the grid

Design reasoning. The one-line principle lives in [`docs/design.md`](../design.md); this file holds the argument and its instances.

## The argument

A response describes what the remote system chose to send, not what exists. Code that treats an absence in the payload as an absence in the world misreports the empty case as the impossible one, and the failure surfaces far from the assumption that caused it. Where a wire format elides the empty case, repair it once at the boundary rather than teaching every consumer to tell the two apart.

## Instances

Sheets returns a row inside the table with no `rowData` when no cell in it holds a value, a formula or a number format, so a sheet left in its designed blank-row state crashed the config sync with a message naming a gid and a column index — the second debugging session that gap has cost. The fix put the missing facts inside the finalize pass that already backfills omitted cells, so "fetched and empty" and "never fetched" stay distinguishable in exactly one place and nowhere else (#17). The same pass already existed because the API omits empty *cells* from rows it does return; the row-shaped and column-shaped versions of that omission are the same fact one axis over.

## Corollaries

The repair has to be narrowed by something authoritative, or it replaces one wrong answer with another. The payload describes every grid column, which on the reported sheet was 23 against a 12-column table, so the table's own range is what says which columns a fact may be about.
