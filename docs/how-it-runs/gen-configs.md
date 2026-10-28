# gen-configs and configs-diff

**`gen-configs` writes to the live spreadsheet before it regenerates the config files; `configs-diff` compares those files with `HEAD`, with no credential and no Sheets call.**

## What gen-configs writes

`sheets-framework gen-configs` **writes** to the live Table Config/Column Config sheets and to business sheets' header rows (adding missing column IDs) before it regenerates the three local config files from the live config sheets, then writes `appConfigs.ts`, which gathers them and augments `Register` (floor vs generated: [`generated-data.md`](../generated-data.md)). It prints the floor report — a one-line summary of what the config-sheet floor created, overwrote or left behind — beside the untyped-columns summary and the declared-cell report, which names any self-describing row whose declared cell it wrote back to the floor seed.

## What configs-diff prints

`sheets-framework configs-diff` prints what changed in those three config files since `HEAD`: Tables, columns and values added, removed or renamed (matched by `tableId` and `columnId`), and each changed field's old and new values, or "No config changes." It is read-only, with no credential and no Sheets call, and `gen-configs` prints the same summary when it finishes, so the summary is how to review a regen without reading the large `columnConfigs.ts`.
