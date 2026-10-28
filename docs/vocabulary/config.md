# Config words

The elaboration behind [docs/vocabulary.md](../vocabulary.md), whose one-line entries are the rules.

- **Config** — the data describing the spreadsheet's own structure, generated from the real spreadsheet rather than freely made up. Lives in each package's `generatedDir`. Naming has three tiers, from a full collection down to a single fact:
  - **`xConfigs`** (plural, e.g. `tableConfigs`) — the whole map, one entry per Table/column/value-name.
  - **`XConfig`** (singular, e.g. `TableConfig`, `ColumnConfig`) — one entry's full record.
  - **trait** — one property picked out of a single config record (e.g. a Table's `sheetGid`), via accessors like `getTableTraitByGid`/`getColumnTraitById`. The same word is reused one layer up for picking a single property out of a `ValueSchema` (`getValTrait`) — "trait" always means "one property of a multi-field record," never a collection.
  - Why the three config-describing Tables' own entries form a fixed floor: [the config-sheet floor](../generated-data/config-sheet-floor.md#the-seed-declares-structure-the-generated-files-hold-identity).
