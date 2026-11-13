# Config words

The elaboration behind [docs/vocabulary.md](../vocabulary.md), whose one-line entries are the rules.

## Config tiers

- **Config** — the data describing the spreadsheet's own structure, generated from the real spreadsheet rather than freely made up. Lives in each package's `generatedDir`. Naming has three tiers, from a full collection down to a single fact:
  - **`xConfigs`** (plural, e.g. `tableConfigs`) — the whole map, one entry per Table/column/value-name.
  - **`XConfig`** (singular, e.g. `TableConfig`, `ColumnConfig`) — one entry's full record.
  - **trait** — one property picked out of a single config record (e.g. a Table's `sheetGid`), via accessors like `getTableTraitByGid`/`getColumnTraitById`. The same word is reused one layer up for picking a single property out of a `ValueSchema` (`getValTrait`) — "trait" always means "one property of a multi-field record," never a collection.
  - **Typed trait**: a column trait the generated configs carry into the types endpoint code compiles against: its value kind (from its column type, a dropdown's options included) and `emptyValueAllowed`. A person changing one makes `tsc` fail where code relied on the old value, so the run report lists each.
  - Why the three config-describing Tables' own entries form a fixed floor: [the config-sheet floor](../generated-data/config-sheet-floor.md#the-seed-declares-structure-the-generated-files-hold-identity).
  - Not "sampled configs": sampled facts are only what a column's top data row gives, and `letApiAccess` and `emptyValueAllowed` are carried forward, or for `emptyValueAllowed` set by an alteration, not sampled.

## Intended state

Planned by [the map to replace Table Config and Column Config](https://github.com/byronbroughten/sheets-framework/issues/178), and named in [its naming ticket](https://github.com/byronbroughten/sheets-framework/issues/187); no code uses these words yet.

- **Intended state**: what the repo says the spreadsheet's structure will be, only for the traits it names; a trait it omits belongs to the sheet. It is the third side of `sync`'s comparison: intended state, the configs, and the live sheet. For now it is `alterations.ts` alone. One trait's value in it is an **intended value**.
  - Not "declarations": TypeScript's declaration files, `declare` and declaration merging would send an agent looking for `.d.ts` files. The verb "declare" stays plain English.
  - Not "desired state": in Kubernetes and Terraform, drift from desired state is reverted. Here a person's change to an intended trait stops the run, and nothing is silently reverted.
  - Nothing names whatever holds an intended value; a message names the pending alterations. Claims, which would add endpoint and chore code to intended state, wait on [#286](https://github.com/byronbroughten/sheets-framework/issues/286).
- **Alteration**: a trait `sync` applies once and then removes, a new column or a deletion included, as in SQL's `ALTER TABLE`. `alterations.ts` holds the **pending alterations** (`pendingAlterations`, built by `makePendingAlterations`) until each lands.
  - Not "one-off": that is a transient chore in `chores/oneOff/`. Not "migration": a migration is kept and recorded as applied, and "migrating" already names moving the config tabs' ticks into intended state. Not "queued": the write queue's queued writes have that word.
- **ID-name map** (`idNameMap.ts`, `IdNameMap`): the generated, type-only file pairing every Table's ID with its name, and every column's for Tables with API access; intended state's keys and `before:` values are typed against it.

## Entered data and held sends

Named in [How does the gate tell a data-preserving gen:configs or chore send from one that destroys data?](https://github.com/byronbroughten/sheets-framework/issues/188), decided in [ADR 0034](../adr/0034-hold-entered-data-loss-at-the-node-hosts-adapter-until-a-hash-bound-confirm.md); no code uses these words yet.

- **Entered data**: a non-empty value in a Table data-row cell that isn't a formula, judged cell by cell. Head-row cells never count, even when a person typed them: their content is in the configs, so git is their undo and a person's edit there is a three-way conflict.
  - Not "business data": the dev spreadsheet's fixture Tables hold none, and the rule covers them. Not "user data": it reads as the person running the command. Not "typed data": the glossary's **Untyped** is about column types. Not "Table data": a Table starts on its header row, so that would include headers.
- **Held**: a send's item that waits for `--confirm <token>`: deleting or overwriting entered data, a forced property, a dropdown style reset under `"ask"`, and anything not classified. A send with no held item is **free**.
  - Not "destructive": a forced property and a dropdown style reset are held but lose no data. Not "gated": that is the tier gate's word, in byro-repo.
