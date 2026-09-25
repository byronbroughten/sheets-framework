# How it runs

Read the heading the task needs.

**The framework runs in two hosts, and only one of them is production**: Apps Script, pushed with `clasp`, and the Node host behind the chore runner and `gen-configs`, which reaches Sheets and nothing else. Every tooling command runs through the `sheets-framework` bin against the spreadsheet in its package's `sheets.config.json` ("The sheets-framework bin").

## Apps Script is the production host

`src/` is TypeScript compiled by `rollup` (via `@rollup/plugin-typescript`, configured by "The rollup preset" below) into a single `dist/bundle.js`, which `clasp` pushes to a Google Apps Script project. The code then executes server-side inside Apps Script. Spreadsheet I/O goes through `GoogleSheetsAPI`, which is the only module that calls the Sheets Advanced Service (`Spreadsheets.get`, `getByDataFilter`, `batchUpdate`), not the `SpreadsheetApp` UI-bound API. Entry points are the top-level functions in an app's entry file (the framework's dev project has its own, `dev/index.ts`) — `triggerOnEdit` and `triggerOnChange`, which Apps Script calls by name from installed triggers — each a one-liner handing its event and the app's configs and endpoints to `AppsScriptApi` (`src/appsScriptHost/`), the Apps Script host's trigger glue. It decodes the event into a platform-neutral `SheetEdit` or `SheetChange` and calls `Api`'s matching static handler, and shows the floor notice the change handler returns, through Google's toast with a ⚠️ title. Before any framework work, the handler installs the configs, which the tiers read through `Register`, and then the `GoogleSheetsAPI` adapter as the `RawSource` in Raw state. The edit handler installs only once the edit looks like an action-row checkbox, so an ordinary edit costs no adapter setup. `GoogleSheetsAPI.forAppsScript()` binds the adapter to the spreadsheet the Apps Script project is bound to, so the project must be container-bound; a standalone script throws asking for that. Node/DOM APIs are not available there — only in the local build tooling (`rollup`, `tsc`, `scripts/`).

## The rollup preset

**`rollup.config.mjs` is `rollupPreset({ input })` from `scripts/rollupPreset.js`**, exported as the `./rollup` subpath, and returns a plain config an app can spread and override. It tree-shakes, so **only functions in the entry file are callable by name from Apps Script**: a function in any other module is dropped when nothing calls it, where the old `treeshake: false` bundle kept it by accident of flattening. The preset keeps every top-level function declaration in the entry file, exported or not, exporting it to rollup and stripping the `export` statement from the chunk. The `es` format's `export` is a syntax error in Apps Script. To opt out, pass `treeshake: false`: `rollupPreset({ input, treeshake: false })` keeps every module's code.

- **`rootDir` widens** (default `.`, resolved from cwd) with `filterRoot: false` and `declaration: false`, because the TypeScript plugin otherwise treats source outside its `rootDir` as external and silently leaves it out of the bundle. An app that bundles the framework from a sibling workspace passes the directory above both.
- **An `UNRESOLVED_IMPORT` warning fails the build**, so a missing import can't ship as an external.
- **Rollup keeps an entry function's own name** when another module declares the same one, and renames the other; the strip throws if it ever sees a renamed entry export.

## The ESLint preset

**An app's `eslint.config.mjs` spreads `@byronbroughten/config`'s `eslintPreset`, then `appEslintPreset({ testSetupFiles })` from `scripts/eslintPreset.js`**, exported as the `./eslint` subpath. It is JavaScript, not TypeScript, because Node won't strip types from a file under `node_modules`. It adds, for the app's `src/`:

- **The platform boundary**: no `GoogleSheets/` import, `GoogleAppsScript` type, or Apps Script global such as `SpreadsheetApp`. The framework's own config applies the same blocks through `sheetsSrcBlocks`, so the boundary can't drift between them.
- **The public entry**: the framework is imported only as `@byronbroughten/sheets-framework`, and `/testing` only from `*.test.ts` and the `testSetupFiles`, which are the non-test files that install configs for tests. `src/generated/**` is exempt, because `gen:configs` writes a relative import of the framework's `makeConfigs`.
- **No relative import out of `src/`**. The `../` count differs by depth, so there is one block per depth, covering `src/` down to five folders deep. A file nested deeper gets none of the import rules; add a depth to the preset before nesting that far.

The preset leaves naming for domain-free utility folders to the app: the app's own config gives its utilities folder `variableNaming`, as the real-estate app does for `src/appUtils/`.

## The Node host

The Node host is the second one (the framework's `src/nodeHost/`, launched by its `scripts/nodeHost.ts`). Spreadsheet I/O is six `RawSource` methods — fetch sheet properties, fetch the time zone, fetch grid ranges, fetch conditional format rules, fetch edit protections, apply the queued write list — none of which names a spreadsheet: the adapter is bound to one when it is constructed. `GoogleSheetsAPI` maps those onto the three Advanced Service verbs and, in Node, an HTTP transport. `NodeHost.ensureGlobals()` installs a `Logger` onto the global scope, installs configs as the Apps Script entry call does, and injects `GoogleSheetsAPI`, bound to the configured spreadsheet ID, as the `RawSource`, before any framework module loads. There is no `PropertiesService` stub. It does **not** install a `Sheets` global, so a chore that reaches for `Sheets` or `SpreadsheetApp` still fails by name. Two commands use it: `sheets-framework chore <name>` and `sheets-framework gen-configs`. A chore gets the four files in its package's `generatedDir`, passed to the same `installConfigs` the entry call uses, so a chore run from the framework package runs on its `dev/generated/`. `gen-configs` installs the package's own configs too, or the framework's dev ones when the package has none yet, since it reads only the config floor, which every package shares.

## The Node host reaches Sheets and nothing else

**The Node host reaches Sheets and nothing else.** Triggers, Gmail and Docs run in Apps Script and stay there — a deliberate boundary, not a gap waiting to be filled. `ScriptApp` and `SpreadsheetApp` are deliberately left uninstalled, so a chore that reaches for one fails by name rather than half-working. And the adapter's fidelity to the real Advanced Sheets Service is an assumption rather than a fact: it is exercised against a recorded payload, not against Google. Moving `gen:configs` off `clasp run` also removed the last routine exercise of the deployed bundle, leaving the live trigger as the only thing that runs it — an accepted cost, taken because one regeneration path beats two, but worth remembering if a deployment-only failure ever appears.

## Node-host transport: one HTTPS call per request

**Each Node-host request is one HTTPS call to the Sheets REST API**, authenticated with the `desktop-clasp-run` credential clasp already stores. The framework's Sheets calls are synchronous and use their return values immediately, so the transport has to block: the framework's `scripts/nodeHost.ts` `spawnSync`s `scripts/fetchSync.js`, a one-request-per-process script, plain JS so each request skips `tsx`, that reads the request from stdin and writes the response to stdout. TypeScript is run by `tsx`, because the repo's relative imports are extensionless under `bundler` module resolution and Node's own type stripping cannot resolve them.

## The sheets-framework bin

**The framework's `scripts/sheets-framework.js` is the one tooling entry**, installed as the `sheets-framework` bin: a JS shim that registers `tsx`, then imports `scripts/cli.ts`, which runs `gen-configs`, `chore`, `probe` or `setup-auth`. The first three read the nearest `sheets.config.json` above cwd, the way clasp finds `.clasp.json`, and take the spreadsheet ID from it and nowhere else: no flag, no env override (`scripts/sheetsConfig.ts`). The file is gitignored and data-only; the package commits a `sheets.config.example.json` to copy from, and the bin fails naming it when only the example exists:

- `spreadsheetId`: the package's spreadsheet.
- `generatedDir`: where `gen-configs` writes the four config files and where the chore runner loads them from.
- `choreHomes`: the package's own chore folders. The framework's generic chores (its `src/chores/`) are listed in every package, and a package chore with a generic chore's name stops the run (`scripts/choreIndex.ts`).

Paths are relative to the config file, which sits at each package's root: the app's has `src/generated` with `src/chores` and `src/chores/oneOff`, the framework's (the dev spreadsheet) `dev/generated` with `dev/chores`. **The bin refuses to run when two `sheets.config.json` files in the repo share a spreadsheet ID**, so a copy-paste mistake can't merge the two targets. `gen-configs` checks its output with the package's own `npm run tsc`.

## The dev project

**The framework's own Apps Script project runs a dev entry, `dev/index.ts`**, over the `Sheets Framework Dev` spreadsheet: the same two one-line trigger globals as an app's, over `devConfigs` and `devEndpoints` (a `runItem_result` fixture endpoint, `dev/devEndpoints/countSelection.ts`). It imports the framework only through `src/framework.ts`. The project's `appsscript.json` is minimal: the Sheets advanced service and the spreadsheets and scriptapp scopes, no Gmail or Docs. The framework package's `build` script is `rollup --config && clasp push`, `push` is `clasp push` alone and `run` is `clasp run`.

- **The dev project needs one one-time human step before a tick runs**: install installable triggers for `triggerOnEdit` (on edit) and `triggerOnChange` (on change). Then tick some rows' Selected box and the Result box in the action row of the `runItem` tab: the Run status column reads `Counted N selected row(s)`.
- **The dev fixture is built in three steps**: `gen-configs` (creates the config floor), `chore buildDevFixtures --send` (creates any missing fixture tab and sets the config ticks, and refuses any other spreadsheet ID), then `gen-configs` again, which writes the framework's checked-in `dev/generated/`. The chore leaves an existing tab alone, so to rebuild a drifted one, delete the tab and run the three again.

## What gen-configs writes

`sheets-framework gen-configs` **writes** to the live Sheet Config/Column Config sheets and to business sheets' header rows (adding missing column IDs) before it regenerates the four local config files from live Spreadsheet Config plus the other config sheets (floor vs generated: [`generated-data.md`](./generated-data.md)). It prints the floor report — a one-line summary of what the config-sheet floor created, overwrote or left behind — beside the untyped-columns summary and the declared-cell report, which names any self-describing row whose declared cell it wrote back to the floor seed.

## The chore and its dry run

A **chore** is a unit of work run from the terminal against the live spreadsheet, as against an endpoint, which an operator runs from the sheet by ticking a checkbox. One typed exported const per file, named after its file, in a generic or package chore home — see [Chores](./architecture/chores.md) for the homes.

```
sheets-framework chore                   # list the chores
sheets-framework chore <name>            # dry run: read the sheet, print what it would write
sheets-framework chore <name> --json     # the same, as raw request JSON
sheets-framework chore <name> --send     # apply it
```

**Dry-run mode is enforced at the adapter, not at the runner and not at the chore's handle.** The adapter is the single door to Google in the Node host, so a dry run that records the requests and returns an empty response makes a *writing* dry run unrepresentable rather than discouraged — whoever calls the flush, the config orchestrator's internal one included. Reads are not gated: a dry run fetches from the live sheet normally, and only sends are suppressed. That is why narrowing a chore's spreadsheet handle to remove the flush would buy nothing and is not done.

The preview is a rendered summary, one line per request, naming the sheet, the range and what changes (`UpdateRequestSummary`). Each capability the framework grows owes the renderer a line format — an accepted recurring cost, because a preview nobody reads converts a decision into a formality. The compiler collects it: `ModeledRequestVerb` in `GoogleSheetsAPI.ts` lists the request kinds the builders may produce, and `UpdateRequestSummary` must format each one or `tsc` fails. `--json` is the escape hatch for when a line looks wrong. Because a dry run withholds every flush, a config regeneration that runs the floor then refetches still sees the live column types, so a generated-file preview can show the value-name churn that `--send` would prevent.

## When the Node host fails to authenticate

`gen-configs` and the chore runner both refresh the named credential `desktop-clasp-run` out of `~/.clasprc.json`. If one fails with an auth error, the token needs re-minting — run `npx sheets-framework setup-auth` from any package (it runs the framework's `scripts/setup-clasp-run-auth.sh`), which walks through it. The consent screen of the GCP project behind that credential should be published to production; left in "Testing" it would issue refresh tokens that expire every 7 days. Publishing alone doesn't fix an existing token, since one minted under "Testing" keeps its expiry — the re-authorization is the part that matters.

Google no longer lets you view or download a client secret after creating it, but you don't need to: clasp stores `client_id` and `client_secret` in `~/.clasprc.json`, and `--creds` reads only those two plus a localhost `redirect_uris` entry, so the file is always rebuildable. The script does that for you. Keep `~/.clasprc.json` at `chmod 600` — the refresh token in it no longer self-expires.

## Seeing the raw Sheets JSON

`GoogleSheetsAPI` maps the payload before anything can log it, and the `gsheets` MCP returns cell values only. To see what Google actually sent, use the committed probe:

```
sheets-framework probe --fields 'sheets(properties(sheetId,title),protectedRanges)'
sheets-framework probe --filter '{"dataFilters":[...]}' [--fields '<mask>']
sheets-framework probe --path sheets.title=Log.protectedRanges      # re-read the last response, no request
```

The bin's `probe` (`scripts/sheetsProbe.ts`) sends one request through the Node host's `SheetsTransport`, authenticated like a chore. That request is a `GET` with a `fields` mask or a `:getByDataFilter`, and the script cannot build any other kind. It writes the full response, pretty-printed, to the gitignored `.probe/last.json` in its package: `.probe/last.json` for the app, `dev/.probe/last.json` for dev. Stdout gets only a summary: top-level keys, array counts, and each sheet's id and title. With `--path`, stdout gets that one subtree, printed whole when it is short and summarized when it is not. When the summary isn't enough, `Read` a line range of the file it names. The throwaway `scripts/*.tmp.mjs` route this replaced is retired.

## The `gsheets` MCP tools

A `gsheets` MCP server can read and write a Google Sheet directly, separately from `clasp`/Apps Script. Take `spreadsheet_id` from the package's `sheets.config.json`.

- **It returns cell values only, so it cannot see a table's declared column types.** `tables[].columnProperties` — a column's `columnType`, its table-column name, its validation rule — is invisible to `get_sheet_data`, and `include_grid_data` reaches cell formats but not tables. Reading those means calling the Sheets REST API with the `clasp` credential, and `probe` (above) is how to do it: it is read-only, like a chore dry run.
