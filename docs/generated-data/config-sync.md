# The config sync and the config sheets

Part of [generated data](../generated-data.md). It keeps a config-sheet cell only for identity or a human, and rewrites existing rows as well as appending.

## Which config-sheet columns get a cell

**A live config-sheet column exists so the host can identify the row (Sheet GID; Column ID on Column Config) or so a human can use it on that sheet** (readable identity such as sheet title and header; declarations such as Empty value allowed, Custom default value, Let api access; helper formulas). Sampled generated traits — `isFormula`, `valueName`, `hasIdColumn`, `idPrefix` — do not earn a cell. Emit still writes those traits into the generated files from the described live columns; leftover Is formula / Value title / Has ID column / ID prefix cells, while those columns still exist, are ignored (#33, #87). A Column Config row is identified by its Sheet GID and Column ID together, never by Column ID alone, because a duplicated tab repeats the original's column IDs (#101).

## The correction pass rewrites existing rows

**The sync doesn't only append rows — it rewrites existing ones.** Its correction pass walks *every* active Column Config row and rewrites that row's sheet title and header from the live tab and header, rows describing the config sheets themselves included — and Sheet Config's sheet title the same way. On a **self-describing row** it also rewrites the declared cell from the floor seed: Let api access on Sheet Config, Empty value allowed on Column Config. It does not rewrite sampled traits onto the config sheets. A header change still propagates on the next run with no hand edit; a value-name or formula-ness change propagates in the generated files because emit samples the live column, not because a staging cell was corrected. The [config-sheet floor](./config-sheet-floor.md) holds because those particular columns don't change, not because the sync skips their rows — and when one of them genuinely does change, the floor moves with it, as it did when Sheet Config's checkbox columns picked up the `checkbox` value name (#12).

## `syncConfigSheetRows` and the sync endpoint

`ConfigCoordinator.syncConfigSheetRows()` runs the config-sheet floor first and flushes it, then bootstraps the live Spreadsheet Config overlay, then syncs the live Sheet Config/Column Config sheets, without writing TypeScript files. It is reachable from the sheet as a `06_API` endpoint (the run flushes after the action returns). `syncAndFlushConfigSheets()` is that same sync plus an in-orchestrator flush, used from tests. From the terminal, `npm run gen:configs` runs that same floor-then-overlay-then-sync path, then regenerates all four files, so there is no separate sync-only command.
