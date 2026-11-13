# sync and configs-diff

**`sync` plans `alterations.ts` against the live sheet and the previous configs, refreshes the config files, and sends the plan only with `--send`; its exit code says whether the run did what it set out to ([exit codes](#syncs-exit-codes)). `configs-diff` compares those files with `HEAD`, with no credential and no Sheets call.**

## What a sync run does

`sheets-framework sync` (`npm run sync` in a package) runs `SyncCoordinator.run` with the pending alterations in `alterations.ts` as intended state and the previous `idNameMap.ts` as the map its entries are checked against; a missing `alterations.ts` is an empty one and a missing map an empty map. A bare run reads the live sheet, plans, and refreshes the configs from what it read, but sends nothing. `--send` sends the plan as well, any missing column IDs to business sheets' header rows, and a Fill Row IDs checkbox into each empty action-row cell above an ID column, in one flush; a non-empty cell there is left alone. After the regen it fills every blank ID cell, planned from the fresh configs, as its own send; that is the only row data `sync` writes. A Table with an ID column but no ID prefix stops the run before any send, and a Table whose ID is its first column gets the `idFirst` notice, since the Table name chip hides its checkbox. Either way, it prints the run's result through `SyncRunReport`.

## sync's entered-data gate and --confirm

**Only a `sync` or chore command holding `--confirm` can lose entered data** ([ADR 0034](../adr/0034-hold-entered-data-loss-at-the-node-hosts-adapter-until-a-hash-bound-confirm.md)). An unstopped dry run flushes its plan after its files are built, for the Node host's gate to judge and withhold; `SyncRunProps.isHostWithholding` turns that flush on, and only the Node host's `sync` sets it. When the gate holds anything, the dry run writes every held item and its value to the package's `.probe/sync-held.json`, prints that path, the held list (each column past 20 lines as counts, with their sum) and the token, then "Rerun with `--send --confirm <token>` to send these". `--send --confirm <token>` reads that file, a missing or unreadable one counting as mismatched, and sends only when the token and the list still match. A refused send writes no file and says only "rerun the dry run": `STOPPED` before any flush landed, `STOPPED PARTWAY` after, with the landed flushes reported.

## What sync writes

A refused run writes no file. Every other run writes every config file, the ID-name map and `appConfigs.ts`, which gathers them and registers them and the map on `Register`, and rewrites `alterations.ts` only when the run returns it, with the landed alterations removed. It leaves the Table Config and Column Config tabs alone. When the `generatedDir`'s parent holds no `alterations.ts`, it creates one in the format sync emits, holding an empty `makePendingAlterations({})`. It prints the configs diff before the report, then runs the package's `npm run tsc`, and fails if that fails.

## sync's exit codes

The run exits 1 when it stopped, was refused, held a flush that stopped a send (before any flush landed or partway), or sent with an item that didn't land (a sent header or Table name Sheets kept another value for); and when the package's `npm run tsc` fails afterward. It exits 0 otherwise: a ready dry run, held items included, or a send where everything landed. `syncExitCode` in `scripts/sync.ts` decides it.

## What configs-diff prints

`sheets-framework configs-diff` prints what changed in those config files since `HEAD`: Tables, tabs, columns and values added, removed or renamed (matched by `tableId`, `sheetGid` and `columnId`), prefixes added to or dropped from the retired list, and each changed field's old and new values, or "No config changes." It is read-only, with no credential and no Sheets call, and `sync` prints the same summary when it finishes, so the summary is how to review a regen without reading the large `columnConfigs.ts`.
