# Config words

The elaboration behind [docs/vocabulary.md](../vocabulary.md), whose one-line entries are the rules.

- **Config** — the data describing the spreadsheet's own structure, generated from the real spreadsheet rather than freely made up. Lives in each package's `generatedDir`. Naming has three tiers, from a full collection down to a single fact:
  - **`xConfigs`** (plural, e.g. `tableConfigs`) — the whole map, one entry per Table/column/value-name.
  - **`XConfig`** (singular, e.g. `TableConfig`, `ColumnConfig`) — one entry's full record.
  - **trait** — one property picked out of a single config record (e.g. a Table's `sheetGid`), via accessors like `getTableTraitByGid`/`getColumnTraitById`. The same word is reused one layer up for picking a single property out of a `ValueSchema` (`getValTrait`) — "trait" always means "one property of a multi-field record," never a collection.
  - Why the three config-describing Tables' own entries form a fixed floor: [the config-sheet floor](../generated-data/config-sheet-floor.md#the-seed-declares-structure-the-generated-files-hold-identity).
  - Not "sampled configs": sampled facts are only what a column's top data row gives, and `letApiAccess` and `emptyValueAllowed` are carried forward, not sampled.

## Intended state

Planned by [the map to replace Table Config and Column Config](https://github.com/byronbroughten/sheets-framework/issues/178), and named in [its naming ticket](https://github.com/byronbroughten/sheets-framework/issues/187); no code uses these words yet.

- **Intended state**: what the repo says the spreadsheet's structure will be, only for the traits it names; a trait it omits belongs to the sheet. It is the third side of `gen:configs`' comparison: intended state, the configs, and the live sheet. One trait's value in it is an **intended value**.
  - Not "declarations": TypeScript's declaration files, `declare` and declaration merging would send an agent looking for `.d.ts` files. The verb "declare" stays plain English.
  - Not "desired state": in Kubernetes and Terraform, drift from desired state is reverted. Here a person's change to an intended trait stops the run, and nothing is silently reverted.
  - Nothing names whatever holds an intended value; a message names the endpoint, the chore, or the pending alterations.
- **Claim**: a trait an endpoint or chore holds in its `claimedTraits`, compared on every run. An agent changing it changes the sheet; a person changing it on the sheet stops the run, which is the sense in which the trait is claimed. "Drop the claim" is removing it.
- **Alteration**: a trait `gen:configs` applies once and then removes, a new column or a deletion included, as in SQL's `ALTER TABLE`. `alterations.ts` holds the **pending alterations** (`pendingAlterations`, built by `makePendingAlterations`) until each lands.
  - Not "one-off": that is a transient chore in `chores/oneOff/`. Not "migration": a migration is kept and recorded as applied, and "migrating" already names moving the config tabs' ticks into intended state. Not "queued": the write queue's queued writes have that word.
- **ID-name map** (`idNameMap.ts`, `IdNameMap`): the generated, type-only file pairing every Table's ID with its name, and every column's for Tables with API access; intended state's `name` fields are typed against it.
