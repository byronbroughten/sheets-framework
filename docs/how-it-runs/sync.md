# sync and configs-diff

**`sync` plans `alterations.ts` against the live sheet and the previous configs, refreshes the config files, and sends the plan only with `--send`; it exits non-zero when the run stops, refuses or doesn't land. `configs-diff` compares those files with `HEAD`, with no credential and no Sheets call.**

## What a sync run does

`sheets-framework sync` (`npm run sync` in a package) runs `SyncCoordinator.run` with the pending alterations in `alterations.ts` as intended state and the previous `idNameMap.ts` as the map its entries are checked against; a missing `alterations.ts` is an empty one and a missing map an empty map. A bare run reads the live sheet, plans, and refreshes the configs from what it read, but sends nothing. `--send` sends the plan as well, and any missing column IDs to business sheets' header rows, in one flush. Either way, it prints the run's result through `SyncRunReport`.

## What sync writes

A refused run writes no file. Every other run writes every config file, the ID-name map and `appConfigs.ts`, which gathers them and registers them and the map on `Register`, and rewrites `alterations.ts` only when the run returns it, with the landed alterations removed. It leaves the Table Config and Column Config tabs alone. When the `generatedDir`'s parent holds no `alterations.ts`, it creates one in the format sync emits, holding an empty `makePendingAlterations({})`. It prints the configs diff before the report, then runs the package's `npm run tsc`, and fails if that fails.

## sync's exit codes

The run exits 1 when it stopped, was refused, held a flush that stopped a send (before any flush landed or partway), or sent with an item that didn't land; and when the package's `npm run tsc` fails afterward. It exits 0 otherwise: a ready dry run, held items included, or a send where everything landed. `syncExitCode` in `scripts/sync.ts` decides it.

## What configs-diff prints

`sheets-framework configs-diff` prints what changed in those config files since `HEAD`: Tables, tabs, columns and values added, removed or renamed (matched by `tableId`, `sheetGid` and `columnId`), prefixes added to or dropped from the retired list, and each changed field's old and new values, or "No config changes." It is read-only, with no credential and no Sheets call, and `sync` prints the same summary when it finishes, so the summary is how to review a regen without reading the large `columnConfigs.ts`.
