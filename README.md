# @byronbroughten/sheets-framework

A TypeScript framework for typed apps on Google Sheets + Apps Script. A live spreadsheet is both the database and the UI: an operator ticks a checkbox in a sheet's action row, and the framework runs the **endpoint** wired to that column, then writes its status back into the rows it was about. The words for what the operator sees are in [`CONTEXT.md`](./CONTEXT.md).

## Install

The package isn't on npm yet. Depend on it from a sibling npm workspace:

```json
"dependencies": { "@byronbroughten/sheets-framework": "*" }
```

It ships TypeScript source, not a build, and your app bundles it. Add the optional peers for what you use: `rollup` and `@rollup/plugin-typescript` to build, `vitest` to test, `eslint` and `@byronbroughten/config` to lint. Pushing to Apps Script also needs [`clasp`](https://github.com/google/clasp).

Your package needs:

- **`tsconfig.json`** extending `@byronbroughten/sheets-framework/tsconfig.base.json`.
- **`rollup.config.mjs`** built from the preset: `export default rollupPreset({ input: "src/index.ts", rootDir: ".." })`, with `rollupPreset` from `@byronbroughten/sheets-framework/rollup`. `rootDir` must cover the framework's source. With tree-shaking on, only the entry file's functions are callable from Apps Script; `treeshake: false` opts out ([the rollup preset](./docs/how-it-runs.md#the-rollup-preset)).
- **`eslint.config.mjs`** spreading `@byronbroughten/config`'s `eslintPreset` and then `appEslintPreset({ testSetupFiles: ["src/installAppConfigs.ts"] })` from `@byronbroughten/sheets-framework/eslint`, which adds the platform and public-entry import rules for your `src/`; `testSetupFiles` are the files besides tests that may import `./testing` ([the ESLint preset](./docs/how-it-runs.md#the-eslint-preset)).
- **`sheets.config.json`** at the package root, naming your spreadsheet ([the bin](#the-bin)). Copy `sheets.config.example.json`; the real file is gitignored.
- **`.clasp.json`** and `appsscript.json` for an Apps Script project bound to your spreadsheet (created from its Extensions menu), with the Sheets advanced service enabled. Copy `.clasp.example.json`; the real file is gitignored. The script reads its spreadsheet from that binding; `sheets.config.json` names it for the Node commands.

## The public entry

App code imports only from `@byronbroughten/sheets-framework`; test files may also import `@byronbroughten/sheets-framework/testing`. Every other path is internal.

**Register your configs once.** `sheets-framework gen-configs` writes four config files into your `generatedDir`. Gather them and augment `Register`, so every Named and endpoint type is typed to your spreadsheet:

```ts
import { columnConfigs } from "./generated/columnConfigs";
import { sheetConfigs } from "./generated/sheetConfigs";
import { spreadsheetConfig } from "./generated/spreadsheetConfig";
import { valueConfigs } from "./generated/valueConfigs";

export const appConfigs = { spreadsheetConfig, sheetConfigs, columnConfigs, valueConfigs };

declare module "@byronbroughten/sheets-framework" {
  interface Register {
    configs: typeof appConfigs;
  }
}
```

An unaugmented `Register` is a type error, not a silent widening.

**Declare the two triggers in your entry file** as one-line globals, and install them as Apps Script triggers (on edit, on change):

```ts
import { Api } from "@byronbroughten/sheets-framework";

const app = { configs: appConfigs, endpoints: myEndpoints };

function triggerOnEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
  Api.handleSheetEdit(app, e);
}

function triggerOnChange(e: GoogleAppsScript.Events.SheetsOnChange) {
  Api.handleSheetChange(app, e);
}
```

`myEndpoints` is an `Endpoints` map keyed by the column whose action-row checkbox triggers each one. An endpoint can't claim a column on the four config sheets, which the framework owns. How a run works: [endpoint dispatch](./docs/architecture/endpoint-dispatch.md).

The entry also exports the Named-tier classes (`SpreadsheetNamed`, `SheetNamed`, `RowNamed` and their bases), the endpoint and run-report types, `Chore` and `SerialDate`. `./testing` exports `stubSheetsService`, `stubLogger`, `EndpointRun`, `installConfigs` and the fake-grid helpers ([testing](./docs/testing.md)).

## The bin

`sheets-framework` runs from any package with a `sheets.config.json` above cwd, and takes the spreadsheet ID from it and nowhere else:

```json
{
  "spreadsheetId": "<your spreadsheet's ID>",
  "generatedDir": "src/generated",
  "choreHomes": ["src/chores"]
}
```

| Command | Does |
| --- | --- |
| `sheets-framework gen-configs` | Ensures the config-sheet floor, then regenerates the four config files from the config sheets. **It writes to the live spreadsheet.** |
| `sheets-framework chore [name] [--send] [--json]` | Lists the chores, dry-runs one (reads live, writes nothing, prints the requests), or applies it with `--send`. |
| `sheets-framework probe --fields\|--filter\|--path …` | One read-only Sheets request; the full JSON goes to `.probe/last.json`. |
| `sheets-framework setup-auth` | Mints the credential the Node host uses, from clasp's login. |

Details: [how it runs](./docs/how-it-runs.md).

## Docs

- [`CONTEXT.md`](./CONTEXT.md): the operator-facing words.
- [`docs/vocabulary.md`](./docs/vocabulary.md): the architecture words.
- [`docs/architecture.md`](./docs/architecture.md): the mechanics.
- [`docs/design.md`](./docs/design.md): why it is shaped this way.
- [`docs/generated-data.md`](./docs/generated-data.md): the generated configs.
- [`docs/how-it-runs.md`](./docs/how-it-runs.md): the hosts and the bin.
- [`docs/testing.md`](./docs/testing.md): the fakes and seams.
- [`docs/style.md`](./docs/style.md): the code-shape rules, layered on the general style doc in `@byronbroughten/config`.

## History

The history from before this package had its own repo is in [byronbroughten/byro-repo](https://github.com/byronbroughten/byro-repo).
