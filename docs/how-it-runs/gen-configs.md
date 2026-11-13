# gen-configs and configs-diff

**`gen-configs` runs sync: it plans `alterations.ts` against the live sheet and the previous configs, sends only with `--send`, and refreshes the config files; `configs-diff` compares those files with `HEAD`, with no credential and no Sheets call.**

## What gen-configs writes

`sheets-framework gen-configs` runs `SyncCoordinator.run` with the pending alterations in `alterations.ts` as intended state and the previous `idNameMap.ts` as the map its entries are checked against; a missing `alterations.ts` is an empty one and a missing map an empty map. A bare run plans and sends nothing. `--send` **writes** the plan and any missing column IDs to business sheets' header rows in one flush. Either way, it prints the run's result through `SyncRunReport`.

A refused run writes no file. Every other run writes every config file, the ID-name map and `appConfigs.ts`, which gathers them and registers them and the map on `Register`, and rewrites `alterations.ts` only when the run returns it, with the landed alterations removed. It leaves the Table Config and Column Config tabs alone. When the `generatedDir`'s parent holds no `alterations.ts`, it creates one in the format sync emits, holding an empty `makePendingAlterations({})`. It prints the configs diff before the report, then runs the package's `npm run tsc`, and fails if that fails.

## What configs-diff prints

`sheets-framework configs-diff` prints what changed in those config files since `HEAD`: Tables, tabs, columns and values added, removed or renamed (matched by `tableId`, `sheetGid` and `columnId`), prefixes added to or dropped from the retired list, and each changed field's old and new values, or "No config changes." It is read-only, with no credential and no Sheets call, and `gen-configs` prints the same summary when it finishes, so the summary is how to review a regen without reading the large `columnConfigs.ts`.
