# The payload is not the grid

Design reasoning. The one-line principle lives in [`docs/design.md`](../design.md); this file holds the argument and its instances.

## The argument

A response describes what the remote system chose to send, not what exists. Code that treats an absence in the payload as an absence in the world misreports the empty case as the impossible one, and the failure surfaces far from the assumption that caused it. Where a wire format elides the empty case, repair it once at the boundary rather than teaching every consumer to tell the two apart.

## Instances

[ADR 0024](../adr/0024-repair-a-payloads-omitted-rows-in-raws-finalize-pass-narrowed-to-the-table.md) repairs a blank row's missing `rowData` inside the finalize pass that already backfills omitted cells, so "fetched and empty" and "never fetched" stay distinguishable in exactly one place, after the gap had cost a second debugging session. The same pass already existed because the API omits empty *cells* from rows it does return, and fills them in for every range fetched that cycle (`TableRaw.finalizeFetches`, `src/02_SpreadsheetRaw/TableRaw.ts`); the row-shaped and column-shaped versions of that omission are the same fact one axis over.

## Corollaries

The repair has to be narrowed by something authoritative, or it replaces one wrong answer with another ([ADR 0024](../adr/0024-repair-a-payloads-omitted-rows-in-raws-finalize-pass-narrowed-to-the-table.md) writes facts only for the Table's columns, since the payload describes every grid column).
