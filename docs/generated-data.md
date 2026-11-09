# Generated data

Each package's `generatedDir` (the app's `src/generated/`, the framework's `dev/generated/`) holds the config files generated from the live spreadsheet, plus `appConfigs.ts`, which gathers them into one `appConfigs` const and registers it and the type-only ID-name map, `idNameMap.ts`, on `Register`. The hand-written `alterations.ts` beside the `generatedDir` is not generated; the regen only creates it empty when missing. The rules for regenerating and editing them are [`src/01_SpreadsheetSchema/AGENTS.md`](../src/01_SpreadsheetSchema/AGENTS.md)'s; what the command writes is [`how-it-runs/gen-configs.md`](./how-it-runs/gen-configs.md#what-gen-configs-writes)'s.

## One file per generated constant, the sync and the floor

| When | File |
| --- | --- |
| The Table Config sheet, `tableConfigs`, `hasIdColumn`, `hasNameColumn`, `idPrefix`, the `tabConfigs` tab list, `retiredIdPrefixes` | [`generated-data/table-configs.md`](./generated-data/table-configs.md) |
| `columnConfigs`, `emptyValueAllowed`, `isFormula`, `valueName` resolution, the untyped-columns summary | [`generated-data/column-configs.md`](./generated-data/column-configs.md) |
| `valueConfigs` and its `const` type parameter | [`generated-data/value-configs.md`](./generated-data/value-configs.md) |
| Which config-sheet columns exist, the correction pass, `syncConfigSheetRows` | [`generated-data/config-sync.md`](./generated-data/config-sync.md) |
| The config-sheet floor, the floor seed, floor identity checks | [`generated-data/config-sheet-floor.md`](./generated-data/config-sheet-floor.md) |

## Reading the generated files by block

**Never read `columnConfigs.ts` whole.** `tableConfigs.ts` and `tabConfigs.ts` are one labeled record per line, and `retiredIdPrefixes.ts` is one line. Grep `columnConfigs.ts` for the Table key (`"occupancy":`) and read that object only: the key opens a multi-line block, and each column config is one labeled line inside it (`columnId`, `header`, `valueName`, `isFormula`, `emptyValueAllowed`, `customDefaultValue`, then the baseline traits `formula`, `columnType`, `actionRowCell`, `groupHeading1`). Same for a single column: grep its `columnId` or name. `valueConfigs.ts` stays a pretty-printed map of member arrays. `idNameMap.ts` is one labeled line per Table and per tab, so grep it for a Table ID, a name or a GID.

## The generated half and the hand-written half

A `generatedDir` holds the generated data files, `tableConfigs.ts`, `tabConfigs.ts`, `retiredIdPrefixes.ts`, `columnConfigs.ts` and `valueConfigs.ts`, plus the generated `appConfigs.ts` and `idNameMap.ts`, and nothing else. Each data file is one literal passed through its validating constructor (`makeTableConfigs`, `makeTabConfigs`, `makeRetiredIdPrefixes`, `makeColumnConfigs`, `makeValueConfigs`), which the framework's main entry, `@byronbroughten/sheets-framework`, exports. The hand-written half sits in the Schema tier's `configReaders/` folder: the generator helpers in `makeConfigs.ts`, the floor seed in `configSheetFloorSeed.ts`, and a sibling types file per constant (`tableConfigsTypes`, `columnConfigsTypes`, `valueConfigsTypes`) holding the derived types and accessor functions built on the data. `gen:configs` writes every data file's import from that entry. The committed `columnConfigs.ts` files keep a relative path into `makeConfigs` until their next regen; that relative path is the one framework deep import the app's lint exempts.

## The ID-name map

`idNameMap.ts` exports one type, `IdNameMap`, and no runtime value. `appConfigs.ts` registers it on `Register`, and the intended-state types in `configReaders/intendedState.ts` read it from there to key `alterations.ts` and each `claimedTraits` by name; until a package has generated its map, they name nothing. It lists every live Table by Table ID with its codebase name (its Table key when it has API access, else its title through the word splitter) and its tab's GID, every tab by GID with its tab name (its title through the word splitter), title and Table IDs, and, for each Table with API access, its columns as column name to column ID. A suite in `scripts/genConfigs.test.ts` asserts the dev package's map and configs agree, so a hand edit or a partial regen fails CI.

## Why the generated files depend on each other

Table keys live as keys in `tableConfigs.ts`, and `columnConfigs.ts` is keyed by those same Table keys; `valueConfigs.ts` in turn depends on `columnConfigs` already being current to know which columns' headers to read. Regenerating a subset after a sheet/column was renamed/added/removed leaves the others referencing stale names, which breaks `npm run tsc` in places that look unrelated (the generated files themselves, plus any hand-written code that names a Table key by string literal).

## What a regeneration runs, on the Node host

Regenerate them all with `sheets-framework gen-configs` (see [`docs/how-it-runs/gen-configs.md`](./how-it-runs/gen-configs.md#what-gen-configs-writes)), which runs `ConfigCoordinator` (`05_Operators`) **on the Node host**, in your working tree; it is the only way to regenerate. It syncs the live Table Config sheet, then the live Column Config sheet (including adding any missing column IDs to business sheets), flushes all of that in one write, then reads the live Value Config sheet, and only then emits source for every config file and the ID-name map. **Live Table sampling on that run — Table header row, column-ID row, first data row — is for this run's Let api access Tables**, after Table Config is loaded, not for every Table. Everyday Table-placement checks still use last-generate sheet GIDs, one regen behind the live box. The command writes every generated file or none, and runs `npm run tsc` itself afterward so a stale hand-written reference surfaces immediately.

## One word splitter derives keys, names and prefixes

**One word splitter, `Str.words`, derives Table keys, column names, value names and new ID prefixes.** A title with whitespace is a sentence, split only on non-letters and non-digits with each word lowercased whole, so `CapEx budget` gives `capexBudget`. A title with no whitespace, like a Table name, is also split on camelCase and acronym boundaries, so `add occ charge`, `Add Occ Charge`, `add_occ_charge` and `addOccCharge` all give `addOccCharge` (prefix `aoc`), and the tab `CapEx` would give `capEx`. A trailing lone `s` stays on its acronym (`unitIDs` gives `unitIds`). `SentenceToCamelCase` mirrors the runtime rule at the type level.

A prefix generated for a new tab splits on punctuation too, so `Item-2B` gives `ib`. A tab whose column ID row already holds IDs keeps the prefix read from it, and a generated prefix skips every retired one ([`retiredIdPrefixes`](./generated-data/table-configs.md#retiredidprefixes-prefixes-no-new-table-may-take)).
