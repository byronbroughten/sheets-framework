# The config sheets and the regen

Part of [generated data](../generated-data.md). The regen reads the live sheet and the previous configs; it neither reads nor writes the Table Config and Column Config tabs.

## Which config-sheet columns get a cell

**A live config-sheet column exists so the host can identify the row (Table ID on Table Config; Table ID and Column ID on Column Config) or so a human can use it on that sheet** (readable identity such as Table name, sheet title and header; declarations such as Empty value allowed and Let api access; helper formulas). Sampled generated traits — `isFormula`, `valueName`, `hasIdColumn`, `idPrefix` — do not earn a cell. Emit writes those traits into the generated files from the described live columns; leftover Is formula / Value title / Has ID column / ID prefix cells, while those columns still exist, are ignored (#33, #87). A Column Config row is identified by its Table ID and Column ID together, never by Column ID alone, because a duplicated tab repeats the original's column IDs (#101).

## The regen leaves the config tabs alone

`ConfigCoordinator.generateConfigFiles()` takes API access from a Table's presence in the previous configs and carries `emptyValueAllowed` forward from them (false for a new column). It writes only missing column IDs onto business sheets' header rows, in one flush, then returns every generated file. There is no sync endpoint, no `syncConfigSheetRows` and no correction pass over the config tabs' rows. The [config-sheet floor](./config-sheet-floor.md) is the `ensureConfigSheetFloor` chore's alone; `npm run gen:configs` does not run it.

A live ID prefix that differs from the configs stops the run before any write (`idPrefixGuard`, run over every Table, sharing the retired prefixes with the title-derived prefix).
