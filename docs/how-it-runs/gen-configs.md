# gen-configs and configs-diff

**`gen-configs` writes business sheets' header rows before it regenerates the config files; `configs-diff` compares those files with `HEAD`, with no credential and no Sheets call.**

## What gen-configs writes

`sheets-framework gen-configs` **writes** to business sheets' header rows (adding missing column IDs) before it regenerates the local config files from the live sheet and the previous configs, then writes `appConfigs.ts`, which gathers them and registers them and the ID-name map, `idNameMap.ts`, on `Register`. It leaves the Table Config and Column Config tabs alone. When the `generatedDir`'s parent holds no `alterations.ts`, it creates one holding an empty `makePendingAlterations({})`, and never touches an existing one. It prints the untyped-columns summary beside the configs diff.

## What configs-diff prints

`sheets-framework configs-diff` prints what changed in those config files since `HEAD`: Tables, tabs, columns and values added, removed or renamed (matched by `tableId`, `sheetGid` and `columnId`), prefixes added to or dropped from the retired list, and each changed field's old and new values, or "No config changes." It is read-only, with no credential and no Sheets call, and `gen-configs` prints the same summary when it finishes, so the summary is how to review a regen without reading the large `columnConfigs.ts`.
