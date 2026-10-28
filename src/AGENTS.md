# Rules for the framework's `src/`

- `00_Source`: the Source: cell values, the host-neutral `RawSource` port, and `GoogleSheets/`, the platform module.
- `01_SpreadsheetSchema`: two halves. The config-free declared layout sits at its root; every reader of the generated configs, which live in each package's `generatedDir`, sits in `configReaders/`.
- `02_SpreadsheetRaw`: positional I/O by row/column index, blind to the configs, which lint holds; a GID reaches a sheet, never a Table, and a Table is reached only by `tableId`; resolves a column only from the Table's live head rows.
- `03_SpreadsheetIdentified`: adds the committed column config, reached by column ID; typed by value name.
- `04_SpreadsheetNamed`: adds names; the end developer's API.
- `05_Operators`: classes on a Named base, suited to one data structure, config regeneration included.
- `06_API`: generic endpoint dispatch (`Api`, `EndpointRun`), handed its endpoint map.
- **Dependencies point only downward, and a file goes in the lowest tier that satisfies it**; `appsScriptHost/`, `chores/`, `nodeHost/` and the two entries sit above every tier.
- **The app imports the framework only from `@byronbroughten/sheets-framework`, and `/testing` only in tests**; lint holds both.
- **Before adding a file, ask "would this make sense in a TypeScript project with no Sheets at all?"** Yes: `@byronbroughten/utils`, never started in this package. Then "in a different Sheets-backed app?" Yes: this package, generically named. No: `packages/real-estate`.
- **`src/` is host-neutral, `nodeHost/` included: no Node or DOM APIs.** It is platform-neutral outside `00_Source/GoogleSheets/`, `appsScriptHost/` and the entry points. Lint holds both.
- **Read by block**: grep `columnConfigs.ts` (`dev/generated/` here) for the Table key (`"occupancy":`) and read that one object; open a long test file's one `describe`.
- **A framework member's callers live in `src/`, `scripts/`, `valueKinds/` and the app's `src/`**; search all four before calling it unused.
- **A guard ships in the same commit as the write it guards, or earlier**: a standing-permission `gen:configs` run can land between any two commits.
- **Read the general style doc (`docs/code-style.md` in `@byronbroughten/config`), then the framework's [docs/code-style.md](../docs/code-style.md), before editing TypeScript here.** Words: [docs/vocabulary.md](../docs/vocabulary.md). Mechanics: the [`docs/architecture.md`](../docs/architecture.md) index.
