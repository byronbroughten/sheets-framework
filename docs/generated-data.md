# Generated data — do not hand-edit

Each package's `generatedDir` (the app's `src/generated/`, the framework's `dev/generated/`) holds four files generated from the live spreadsheet. Regenerate all four together with `sheets-framework gen-configs`, never hand-edit them, and read `columnConfigs.ts` by block. The rules are [`src/01_SpreadsheetSchema/AGENTS.md`](../src/01_SpreadsheetSchema/AGENTS.md)'s; what the command writes is [`how-it-runs.md`](./how-it-runs.md#what-gen-configs-writes)'s.

## One file per generated constant, the sync and the floor

| When | File |
| --- | --- |
| `sheetConfigs`, `hasIdColumn`, `hasNameColumn`, `idPrefix` | [`generated-data/sheet-configs.md`](./generated-data/sheet-configs.md) |
| `columnConfigs`, `emptyValueAllowed`, `isFormula`, `valueName` resolution, the untyped-columns summary | [`generated-data/column-configs.md`](./generated-data/column-configs.md) |
| `valueConfigs` and its `const` type parameter | [`generated-data/value-configs.md`](./generated-data/value-configs.md) |
| `spreadsheetConfig`, layout values | [`generated-data/spreadsheet-config.md`](./generated-data/spreadsheet-config.md) |
| Which config-sheet columns exist, the correction pass, `syncConfigSheetRows` | [`generated-data/config-sync.md`](./generated-data/config-sync.md) |
| The config-sheet floor, the floor seed, floor identity checks | [`generated-data/config-sheet-floor.md`](./generated-data/config-sheet-floor.md) |

## Reading the generated files by block

**Never read `columnConfigs.ts` whole.** `sheetConfigs.ts` is the sheet list — one labeled sheet record per line. Grep `columnConfigs.ts` for the sheet key (`"occupancy":`) and read that object only: the key opens a multi-line block, and each column config is one labeled line inside it (`columnId`, `header`, `valueName`, `isFormula`, `emptyValueAllowed`, `customDefaultValue`). Same for a single column: grep its `columnId` or name. `valueConfigs.ts` stays a pretty-printed map of member arrays.

## The generated half and the hand-written half

A `generatedDir` holds the four generated data files and nothing else: `spreadsheetConfig.ts`, `sheetConfigs.ts`, `columnConfigs.ts` and `valueConfigs.ts`. Each is one literal passed through its validating constructor (`makeSpreadsheetConfig`, `makeSheetConfigs`, `makeColumnConfigs`, `makeValueConfigs`), imported from `../makeConfigs`. The hand-written half sits one level up at the Schema tier root: the generator helpers in `makeConfigs.ts`, the floor seed in `configSheetFloorSeed.ts`, and a sibling types file per constant (`spreadsheetConfigTypes`, `sheetConfigsTypes`, `columnConfigsTypes`, `valueConfigsTypes`) holding the derived types and accessor functions built on the data. A generated file imports `makeConfigs` from the framework by the relative path `gen:configs` computes, the one framework deep import the app's lint exempts.

All four are (or are meant to be) mechanically generated from the real spreadsheet, not hand-authored.

## Regenerate all four together, never a subset

The rules on regenerating and on tab spelling: [`src/01_SpreadsheetSchema/AGENTS.md`](../src/01_SpreadsheetSchema/AGENTS.md).

**`spreadsheetConfig`, `sheetConfigs`, `columnConfigs`, and `valueConfigs` must always be regenerated together, in the same run — never a subset of them.** Sheet names live as keys in `sheetConfigs.ts`, and `columnConfigs.ts` is keyed by those same names; `valueConfigs.ts` in turn depends on `columnConfigs` already being current to know which columns' headers to read. Regenerating a subset after a sheet/column was renamed/added/removed leaves the others referencing stale names, which breaks `npm run tsc` in places that look unrelated (the generated files themselves, plus any hand-written code — like `SheetNameGroups.ts` — that references a sheet name by string literal).

## What a regeneration runs, on the Node host

Regenerate all four with `sheets-framework gen-configs` (see [`docs/how-it-runs.md`](./how-it-runs.md#what-gen-configs-writes)), which runs `ConfigCoordinator` (`05_Operators`) **on the Node host**: it reads live Spreadsheet Config and overlays that record on `ssConfigGet` for the rest of the run, syncs the live Sheet Config sheet, then the live Column Config sheet (including adding any missing column IDs to business sheets), flushes all of that in one write, then reads the live Value Config sheet, and only then emits source for all four files. **Live Table sampling on that run — Table header row, column-ID row, first data row — is for this run's Let api access sheets**, after Sheet Config is loaded, not for every tab. Everyday Table-placement and extra-Table checks still use last-generate sheet GIDs, one regen behind the live box. The command writes all four files or none, and runs `npm run tsc` itself afterward so a stale hand-written reference surfaces immediately.

## The `clasp run` path is gone

**There is exactly one way to regenerate, and the `clasp run` path is gone.** The command used to ask Google to execute `generateConfigFiles` inside the Apps Script project, which holds whatever was last pushed — so editing a config operator locally left `gen:configs` silently running the previously pushed generator, and a generator failure came back as a JSON error blob rather than a stack trace. It now runs the generator in your working tree. Keeping both paths would create two regenerations that can disagree about what the config sheets say, which is the hazard the never-run-a-partial-regeneration rule exists to prevent.

## A tab title becomes the sheet's key

**A sheet's key in `sheetConfigs`/`columnConfigs` is derived from its tab title**, so a misspelled tab becomes a misspelled identifier in the generated files and in every string literal naming that sheet. Fix a tab's spelling before code references it; afterwards it costs a sheet edit, a regeneration and every call site.

## Never hand-edit the data; fix the sheet and regenerate

Do not hand-edit (or AI-edit) the literal data inside `spreadsheetConfig`/`sheetConfigs`/`columnConfigs`/`valueConfigs` once they have real generators — always regenerate from the spreadsheet instead. Structural/type changes around them (not the data itself) are fine. A sheet-shape bug is fixed on the sheet, then regenerated; after any sheet change, read the regenerated entry before building on it.
