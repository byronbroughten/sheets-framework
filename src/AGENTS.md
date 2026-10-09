# Rules for the framework's `src/`

- `00_Source`: the Source: cell values, the host-neutral `RawSource` port, and `GoogleSheets/`, the platform module.
- `01_SpreadsheetSchema`: the declared layout plus every reader of the generated configs, which live in each package's `generatedDir`.
- `02_SpreadsheetRaw`: positional I/O by sheet GID and row/column index, blind to column config; resolves a column only from the Table's live head rows.
- `03_SpreadsheetIdentified`: adds the committed column config, reached by column ID; typed by value name.
- `04_SpreadsheetNamed`: adds names; the end developer's API.
- `05_Operators`: classes on a Named base, suited to one data structure, config regeneration included.
- `06_API`: generic endpoint dispatch (`Api`, `EndpointRun`), handed its endpoint map.
- **Dependencies point only downward; lint holds the numbered tiers to it.** A file goes in the lowest tier that satisfies it. Utilities are the external `@byronbroughten/utils` package, imported from `@byronbroughten/utils/<module>`; `appsScriptHost/`, `chores/`, `nodeHost/` and the two entries sit above them all. **The app imports the framework only from `@byronbroughten/sheets-framework` (`framework.ts`), and `/testing` (`frameworkTesting.ts`) only in `*.test.ts` and its config setup file**; lint holds it.
- **Before adding a file, ask "would this make sense in a TypeScript project with no Sheets at all?"** Yes: `@byronbroughten/utils`, never started in this package. Then "in a different Sheets-backed app?" Yes: this package, generically named. No: `packages/real-estate`.
- **`src/` is host-neutral, `nodeHost/` included: no Node or DOM APIs.** It is platform-neutral outside `00_Source/GoogleSheets/`, `appsScriptHost/` and the entry points. Lint holds both.
- **Regenerate, never hand-edit, a package's generated data** (`dev/generated/` here). Read by block: grep `columnConfigs.ts` for the sheet key (`"occupancy":`) and read that one object; open a long test file's one `describe`.
- **A guard ships in the same commit as the write it guards, or earlier**: a standing-permission `gen:configs` run can land between any two commits.
- **Read the general style doc (`docs/code-style.md` in `@byronbroughten/config`), then the framework's [docs/code-style.md](../docs/code-style.md), before editing TypeScript here.** Words: [docs/vocabulary.md](../docs/vocabulary.md). Mechanics: the [`docs/architecture.md`](../docs/architecture.md) index.
