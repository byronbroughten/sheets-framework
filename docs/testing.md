# Testing

Read the heading the task needs.

Vitest, as plain Node against fakes, so `npm test` is always safe. The test-writing rules are the house style's Tests section; this file is the fakes, the seams and the exemplar columns. Most tests need only `stubSheetsService()`.

## Running the tests

Tests run on [Vitest](https://vitest.dev): `npm test` (single run), `npm run test:watch`, or `npm run test:coverage`, from the framework's own `vitest.config.ts`. None of it touches the live spreadsheet or Apps Script — it's plain Node against fakes — so it's always safe to run freely, same as `npm run tsc`.

## Co-located tests and tier imports

A test is co-located with what it tests — `Foo.ts` → `Foo.test.ts` in the same folder. This obeys the same downward-only tier-import rule as production code: a test only imports from its own tier and below.

## Faking spreadsheet I/O with `stubSheetsService`

This is not a Node app at runtime (see [`docs/how-it-runs.md`](./how-it-runs.md)). Tests that construct `SpreadsheetRaw` inject a `RawSource` via `stubSheetsService()`, which builds a `GoogleSheetsAPI` on a fixture-backed Advanced Service lookalike — they do not install `globalThis.Sheets`. `src/testSupport/` holds shared fakes, sitting outside the numbered tiers (like `utils/`) so any tier's tests can import it:

## `fakeSheetsService.ts`: the fixture-backed Sheets service

`stubSheetsService()` installs that adapter, bound to a fixed spreadsheet ID, so no test stubs a bound spreadsheet to supply one. Fixture `get` / `getByDataFilter` are backed by a real in-memory grid and typed against the actual `GoogleAppsScript.Sheets.Schema` types, so a fixture that drifts from the real response shape is a compile error.

## Which requests the fake replays

`batchUpdate` is mostly a spy — it records the exact Google requests the adapter mapped but doesn't replay them onto the fixture (`addSheet` and `addTable` are recorded and never applied, so a fixture never gains the tab or Table, and the `updateTable` that follows an `addTable` finds no fixture Table to replay onto); the Sheets request grammar (`appendCells`/`updateCells`/`insertDimension`/`sortRange`/...) is large, so extend this as tests come to need particular request kinds applied back. The two conditional-format request kinds are the exception: a fixture may list `conditionalFormats`; `get` returns a non-empty list and `getByDataFilter` never does, as the live API behaves, and an add inserts at its index while a delete removes and renumbers, so ordering tests assert the list a flush left rather than only the recorded indexes. Edit-protection adds and deletes are replayed onto `protectedRanges` the same way. An `updateTable` with `fields: "columnProperties"` replaces the fixture Table's whole column list, as the live API does: types and validation rules not sent are dropped, each sent `columnName` is written to the header cell, and a column with no `columnName` is rejected. So a second floor ensure after a flush sees the types it just set.

## `getByDataFilter` ignores its filter

`get` records its params in `getCalls`, so a test can count property fetches. `getByDataFilter` ignores the filter it is handed: it records the resource for assertions, then returns every fixture sheet in full, as `get` does apart from conditional formats. A test can therefore assert *what was requested* but never *what a filter would have returned*, and any test reasoning about a partial response is vacuous. The real API's filter-dependent behavior is a live source of bugs the fake cannot reproduce, most notably that a sheet's `tables` metadata comes back only for a filter whose range overlaps the table.

## Fixture Tables, extra Tables and absent rows

A fixture's table starts where `spreadsheetConfig` says it must, so a fixture never asserts a layout production would refuse; `table.startRowIndex`/`startColumnIndex` override that for a deliberately misplaced table, and `table.endColumnIndex` narrows the table inside a wider grid, which is what the real payload looks like whenever a sheet has columns past its table. A fixture lists extra Tables in `extraTables` beside that single `table` so a test can emit more than one without teaching the fake filter-range arithmetic. The two absent-row hooks model the two shapes a row with nothing in it really comes back as: `rowsWithNoGridData` emits a block describing every column and carrying no `rowData`, which is what a blank row inside the grid produces, and `rowsWithNoGridBlock` emits no block at all, which is what a row past the populated grid produces (#17). The one filter-aware behaviour is a per-sheet `isTableHiddenFromFilteredFetch` flag, which withholds that sheet's every Table from `getByDataFilter` while `get` still returns them — a deliberate escape hatch for the probe blind spot described in [round trips](./architecture/round-trips.md), not real range arithmetic, which would put a second and subtly wrong model of the Sheets API into test support.

## `fakeAppsScriptGlobals.ts` and `fakeSheetConfigSheet.ts`

`fakeAppsScriptGlobals.ts` — `stubScriptAndSpreadsheetApp()` (a fluent trigger builder covering `AppsScript.trigger`'s usage, and an active spreadsheet whose `getId` feeds `GoogleSheetsAPI.forAppsScript()` and whose `toast` records each message with its title and timeout; `{ spreadsheetId: null }` stands in for a standalone script).

`fakeSheetConfigSheet.ts` — a fake "Sheet Config" sheet, for behaviour that reads or writes a whole row rather than one named cell (clearing, the blank test, the wipe, append reuse). Anything working from a sheet's *configured* columns resolves each of them against the live columnId row, so such a test needs a fixture listing every column the config declares — which makes Sheet Config the right subject, as the smallest sheet both config sets share. Its GID and column IDs come from the installed configs, so it serves either test program.

## Three seams

SpreadsheetRaw / EndpointRun tests inject a RawSource (`stubSheetsService`). `GoogleSheetsAPI.test.ts` asserts mapping: local operations in, Google requests out; Google payload in, Raw-facing snapshot out. The Node host's HTTP transport is tested at the wire in that same file and in `NodeHost.test.ts`: with a dry run armed, a real framework write through `SpreadsheetRaw` puts nothing on the wire while a read still reaches it. That is the whole safety property in one assertion, and it is the most valuable test around the Node host. The preview renderer is a pure function — Google requests in, lines out — and is tested directly with no seam. The raw request opening is tested at the existing seam, since the fixture already records mapped batch updates and can show a raw request landing last. Chores are not tested (see [Chores](./architecture/chores.md)), and the chore runner's name lookup stays thin enough not to warrant a test — revisit that judgement if it grows past resolving a name to a module.

## The Apps Script surface and its fakes

Schema/config resolution and ID encode/decode need no mocking. The GAS-touching surface is narrow: `00_Source/GoogleSheets/AppsScript.ts`, `GoogleSheetsAPI.forAppsScript()`, and `appsScriptHost/AppsScriptApi.ts`, which decodes the trigger events with `AppsScript.sheetEdit` / `AppsScript.sheetChange` and is tested against faked events and globals. Production Raw does not read `Sheets` from global scope. A change that first reaches a new Apps Script global adds its wrapper under `00_Source/GoogleSheets/` and extends `fakeAppsScriptGlobals.ts` in the same change (the rule: [`src/00_Source/GoogleSheets/AGENTS.md`](../src/00_Source/GoogleSheets/AGENTS.md)).

## Two test programs, two config sets

The framework tests (tiers `00`–`06`, `appsScriptHost/`, `nodeHost/`, `testSupport/` and `utils/`) run on the dev spreadsheet's configs, and the app's tests, today `businessEndpoints/`, run on the app's. Each package is its own Vitest project and `tsc` program, so each program carries one `Register` augmentation: the framework's `dev/devConfigs.ts` and the app's `src/appConfigs.ts`, each installed for its tests by `dev/installDevConfigs.ts` and `src/installAppConfigs.ts`. The app's program reaches framework source only through imports, so framework tests never enter it, and `testSupport/` compiles in both. An app test takes its fakes and `EndpointRun` only from `@byronbroughten/sheets-framework/testing`. A framework test reads a GID or column ID through `getSheetTraitByName`/`getColumnTraitByName`, and a config sheet's columns through `installedConfigs()`, never by importing a `generated/` file. It never names a real-estate sheet or column, and a made-up sheet in a fixture takes a neutral name such as `Widget`.

## The dev fixtures and their exemplar columns

A framework test names a sheet or column of the dev spreadsheet's fixtures, never a real-estate one. `buildDevFixtures` is their checked-in recipe, and the framework's `dev/generated/` is what `gen-configs` read back from them:

| Sheet | Shape | Exemplars |
| --- | --- | --- |
| `item` | ID and Name columns | The main subject of the Raw and Identified tests. `optionalNote` has Empty value allowed ticked, and `requiredCount` has it unticked. |
| `valueTypes` | ID column, one column per framework value name | `sampledBoolean` only *samples* as boolean, with no column type declared. `checkbox` is a declared checkbox. `dateValue` is a declared date. |
| `log` | No ID column | |
| `runItem` | Endpoint sheet: `selected` checkbox, `result` entry column, `startTime` and `runStatus` | The subject of the `EndpointRun`, `Api`, `Endpoints` and `AppsScriptApi` tests, and of the tier 04 tests that need a checkbox column or a conditional-format or protection fixture. |
| `computed` | No ID column | `rowNumber` is the one formula column. |
| `dates` | ID column, two date columns | `requiredDate` has Empty value allowed unticked and `optionalDate` has it ticked, the only ticked column that is not text, so a ticked read's `SerialDate \| ""` type has a subject. |

Add a fixture sheet only when a test needs a config shape these don't cover. A layout edge case, such as a blank row or a misplaced Table, stays row data in a fake-service fixture. The dev value configs are empty, so no framework test names a dropdown value name.

An exemplar's value name has to stay put. `sampledBoolean` must not be formatted as a checkbox in Sheets: that would regenerate it as `checkbox` and break `SpreadsheetSchema.test.ts`. The same holds for the Empty value allowed ticks: untick `optionalNote` and a test of the ticked case quietly proves the unticked one instead of failing. To change a fixture, edit the recipe, delete the tab and rebuild it ([The dev project](./how-it-runs.md#the-dev-project)).

## Testing an endpoint through `EndpointRun`

An endpoint is tested through `EndpointRun`, never by calling its action (the rule: the house style's Tests section): the seam is the run's entry point, driven by the fake Sheets service, and the assertion is the batch-update requests the run emits — which cells were written, with what values, in what order. Going through the run is what buys the selector pruning, the setup flush and the interplay between a wipe and the appends that follow it, all of which a rebuild-from-scratch endpoint depends on; the cost is a larger fixture, which is the right trade. `businessEndpoints/buildLedger.test.ts` stubs six sheets at once and decodes the recorded `updateCells` requests back into a table of ledger rows, last write per cell winning, since requests apply in order. No test reaches for a private helper, a comparator or an intermediate list of lines. The framework half of an endpoint's behaviour is tested separately in `06_API/EndpointRun.test.ts`, against synthetic endpoints on the `runItem` fixture.

## The fake answers what the adapter asked; Google doesn't

Live JSON omits empty lists and zero-valued fields (a gid-0 sheet, column A, row 1), echoes each colour as a `*ColorStyle`, and `getByDataFilter` drops sheet-level fields such as `conditionalFormats`. A new read isn't done until a chore dry run has read it from the live sheet (the rule: [`src/00_Source/GoogleSheets/AGENTS.md`](../src/00_Source/GoogleSheets/AGENTS.md)).

## Live-sheet verification

There's no standing integration-test tier against a real spreadsheet. When Claude is asked to extend this test infrastructure, it may use the `gsheets` MCP ad hoc (within the repo's gsheets write rules; what the MCP can see is [`docs/how-it-runs.md`](./how-it-runs.md#the-gsheets-mcp-tools)'s) to sanity-check that a fake's behavior actually matches the real API — that stays a manual verification step, never part of `npm test`/CI.

## Navigation assertions

Each of `02`–`04` has a spreadsheet-level test pinning every Meta/primary navigation edge — `sheet`/`sheetMeta`, each view's `column` and `sheet`, the `meta` and `primary` crossings, and the primary sheet's row accessors — with an identity-based type assertion *and* an `instanceof` check. The type assertion catches a declaration that widened; the instance check catches an accessor declared for one view but wired to the other's constructor, which is structurally valid TypeScript and would otherwise surface only against the live spreadsheet.

## CI

The package's own workflow is [`.github/workflows/ci.yml`](../.github/workflows/ci.yml). Coverage is reported, not gated — no failure threshold yet.

## Type verification is `npm run tsc`

`npm run tsc` remains the whole *type*-verification story — run it (and the tests) before considering any change complete, and treat any new type error as something you introduced unless you've confirmed otherwise (check whether the same error exists on a clean checkout, or ask).
