# The Apps Script host and its presets

**Apps Script is the production host**: `rollup` bundles `src/` with the framework's rollup preset, `clasp` pushes the bundle, and an app's ESLint preset keeps its `src/` inside the platform boundary.

## Apps Script is the production host

`src/` is TypeScript compiled by `rollup` (via `@rollup/plugin-typescript`, configured by "The rollup preset" below) into a single `dist/bundle.js`, which `clasp` pushes to a Google Apps Script project. The code then executes server-side inside Apps Script. Spreadsheet I/O goes through `GoogleSheetsAPI`, which is the only module that calls the Sheets Advanced Service (`Spreadsheets.get`, `getByDataFilter`, `batchUpdate`), not the `SpreadsheetApp` UI-bound API. Entry points are the top-level functions in an app's entry file (the framework's dev project has its own, `dev/index.ts`) — `triggerOnEdit`, which Apps Script calls by name from an installed trigger — a one-liner handing its event and the app's configs and endpoints to `AppsScriptApi` (`src/appsScriptHost/`), the Apps Script host's trigger glue. It decodes the event into a platform-neutral `SheetEdit` and calls `Api.handleSheetEdit`. Before any framework work, the handler installs the configs, which the tiers read through `Register`, and then the `GoogleSheetsAPI` adapter as the `RawSource` in Raw state. The edit handler installs only once the edit is a checkbox on the row just above a managed Table's recorded header, inside its recorded columns, a match it reads from the installed `tableConfigs` without a fetch, so an ordinary edit, a data checkbox included, costs no adapter setup. `GoogleSheetsAPI.forAppsScript()` binds the adapter to the spreadsheet the Apps Script project is bound to, so the project must be container-bound; a standalone script throws asking for that. Node/DOM APIs are not available there — only in the local build tooling (`rollup`, `tsc`, `scripts/`).

## The rollup preset

**`rollup.config.mjs` is `rollupPreset({ input })` from `scripts/rollupPreset.js`**, exported as the `./rollup` subpath, and returns a plain config an app can spread and override. It tree-shakes, so **only functions in the entry file are callable by name from Apps Script**: a function in any other module is dropped when nothing calls it, where the old `treeshake: false` bundle kept it by accident of flattening. The preset keeps every top-level function declaration in the entry file, exported or not, exporting it to rollup and stripping the `export` statement from the chunk. The `es` format's `export` is a syntax error in Apps Script. To opt out, pass `treeshake: false`: `rollupPreset({ input, treeshake: false })` keeps every module's code.

- **`rootDir` widens** (default `.`, resolved from cwd) with `filterRoot: false` and `declaration: false`, because the TypeScript plugin otherwise treats source outside its `rootDir` as external and silently leaves it out of the bundle. An app that bundles the framework from a sibling workspace passes the directory above both.
- **An `UNRESOLVED_IMPORT` warning fails the build**, so a missing import can't ship as an external.
- **Rollup keeps an entry function's own name** when another module declares the same one, and renames the other; the strip throws if it ever sees a renamed entry export.

## The ESLint preset

**An app's `eslint.config.mjs` spreads `@byronbroughten/config`'s `eslintPreset`, then `appEslintPreset({ testSetupFiles })` from `scripts/eslintPreset.js`**, exported as the `./eslint` subpath. It is JavaScript, not TypeScript, because Node won't strip types from a file under `node_modules`. It adds, for the app's `src/`:

- **The platform boundary**: no `GoogleSheets/` import, `GoogleAppsScript` type, or Apps Script global such as `SpreadsheetApp`. The framework's own config applies the same blocks through `sheetsSrcBlocks`, so the boundary can't drift between them.
- **The public entry**: the framework is imported only as `@byronbroughten/sheets-framework`, and `/testing` only from `*.test.ts` and the `testSetupFiles`, which are the non-test files that install configs for tests. `src/generated/**` is exempt, because `gen:configs` writes a relative import of the framework's `makeConfigs`.
- **The generated boundary**: outside `src/generated/`, a module imports only `generated/appConfigs`, never one of the three config files it gathers. Files inside `generated/` import each other freely.
- **No relative import out of `src/`**. The `../` count differs by depth, so there is one block per depth, covering `src/` down to five folders deep. A file nested deeper gets none of the import rules; add a depth to the preset before nesting that far.
