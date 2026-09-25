# Config words

The elaboration behind [docs/vocabulary.md](../vocabulary.md), whose one-line entries are the rules. These are the **architecture** words; the operator-facing ones live in [CONTEXT.md](../../CONTEXT.md).

If you're renaming or relocating something and aren't sure which word fits, ask rather than guess: these distinctions were deliberately hashed out, and folder placement depends on them.

- **Config** — the data describing the spreadsheet's own structure, generated from the real spreadsheet rather than freely made up. Lives in each package's `generatedDir`. Naming has three tiers, from a full collection down to a single fact:
  - **`xConfigs`** (plural, e.g. `sheetConfigs`) — the whole map, one entry per sheet/column/value-name.
  - **`XConfig`** (singular, e.g. `SheetConfig`, `ColumnConfig`) — one entry's full record.
  - **trait** — one property picked out of a single config record (e.g. a sheet's `sheetGid`), via accessors like `getSheetTraitByGid`/`getColumnTraitById`. The same word is reused one layer up for picking a single property out of a `ValueSchema` (`getValTrait`) — "trait" always means "one property of a multi-field record," never a collection.
  - `spreadsheetConfig` is the one exception to the plural/singular split: there's only one spreadsheet, so it's a single record with no separate `spreadsheetConfigs` collection.
  - `sheetConfigs`/`columnConfigs`'s own entries for the four config-describing sheets — `sheetConfig`, `columnConfig`, `spreadsheetConfig`, and `valueConfig` — are guaranteed to always come out the same on regeneration, forming a fixed floor underneath the generation process itself (it needs a fixed way to find those sheets and their own columns before it can generate anything else). Treat that floor as guaranteed, not as data to "fix" by trying to regenerate it away.
