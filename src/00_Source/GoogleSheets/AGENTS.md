# Rules for `src/00_Source/GoogleSheets/`

- **This folder, `src/appsScriptHost/` and the entry points are the only code allowed to know the platform**: `GoogleAppsScript.*` types, `SpreadsheetApp`, the Sheets service.
- **Everything outside reaches it through the `RawSource` port** in `../RawSource/`. Translate here, and hand back platform-neutral types.
- **It still touches no Node or DOM API**: both hosts run it.
- **A change that first reaches a new Apps Script global adds its wrapper here** and extends `src/testSupport/fakeAppsScriptGlobals.ts` in the same change.
- **A new read isn't done until a chore dry run has read it from the live sheet**: the fake answers what the adapter asked, and Google doesn't.
