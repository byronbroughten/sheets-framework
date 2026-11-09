# Type-check cost

Map fragment. Sibling headings live in this folder.

How to measure type-check cost, the baseline, the instantiation budget, the one known cliff, and what each typed change measured when it landed.

## How to measure

The generated config describes ~574 columns across ~28 sheets, and a lot of this codebase's typing is mapped/conditional types over those unions. That work is cheap until it isn't, so measure rather than guess: `npx tsc --noEmit --extendedDiagnostics` is always safe to run and reports instantiations and check time.

**Take both readings from clean checkouts of the two commits** — `git worktree add <dir> <commit>`, with `node_modules` symlinked in, is the quick way. Instantiation counts are deterministic for a given TypeScript version, so two runs of the same commit that disagree mean the trees differed, not that the number is noisy; a reading taken from a working tree mid-change is not comparable to one taken after the commit. Check times *are* noisy, so compare them only within one measurement run. A wrong "before" number is the easy mistake here, and it silently turns a saving into a cost.

## Baseline and the lazy mapped filter

The baseline, and what one mapped filter over the flat column map (`{ [K in ColumnFullName]: … }`) costs per distinct instantiation, are [measured in the Estimates](https://github.com/byronbroughten/byro-repo/blob/master/docs/agents/codebase-priorities/estimates.md#type-check). The filter is cheap, and cheaper *lazily* than precomputed: eagerly grouping every value name up front cost ~390k, while filtering per use pays one filter's cost for each value name actually used.

## The known cliff: a widened template literal over both names

**The known cliff is a widened template literal over both names.** Building `` `${TN}${CodebaseNameDelimiter}${CN}` `` when both are unions enumerates the full ~28 × ~574 cross product before any intersection can prune it, which produces `TS2590: Expression produces a union type that is too complex to represent` and takes check time to **~7s**. Going the other way — from a full name to its parts, by indexed access on the flat map — costs nothing. That asymmetry is the measurement behind the missing bridge in [Two ways to address a column](./column-addressing.md).

## Measurements after each change

### #5 endpoint entry

After the endpoint entry landed (#5, Sept 2026): **~620k instantiations, ~1.54s check time**, against **583k, ~1.58s** immediately before it on the same machine — so the entry shape cost ~37k instantiations and no check time. Most of that is the two `ColumnNameFiltered` filters instantiated at the *widened* Table name, which the dispatch boundary forces; the correlated `Endpoints` map itself measured ~200. Composing the two filters (`ColumnValueName<TN, FeedbackColumnName<TN>>`, to re-derive a feedback column's value type inside `EndpointRun`) cost a further ~43k on its own and was dropped — the entry already pins the value type, so the run resolves the column by `columnId` through the Identified tier instead.

### Absolute addressing family

After the absolute addressing family landed (Sept 2026): **~597k instantiations, ~1.6s check time**, against **286k, ~1.5s** re-measured on the same machine in the same session — so instantiations roughly doubled while check time barely moved. (Compare check times only within one measurement run: the Estimates' 1.2s baseline and the 1.5s here are the same baseline commit on different machines.) Most of the increase is the handful of distinct `ColumnFullName<VN, IF>` and `ColumnNameFiltered<TN, VN, IF>` instantiations the endpoint unions and the checkbox operator ask for.

### #11 selector object

After the selector became an object (#11, Sept 2026): **~583k instantiations, ~1.63s check time**, against **638k, ~1.72s** immediately before it on the same machine — so bundling the retain flag into the selector cost nothing and measured ~55k *fewer* instantiations. The saving wasn't chased down; what matters for the budget is that the reshape adds none.

### #12 `checkbox` value name

After the `checkbox` value name landed (#12, Sept 2026): **~610k instantiations, ~1.6s check time**, against **585k, ~1.61s** immediately before it on the same machine — so the new value name and its blank exception cost ~25k instantiations and no check time. Two things paid for themselves: correcting Raw's `OrEmpty` return types to declare the blank *reduced* instantiations, and replacing `StrictExclude<V, "">` with `NotEmpty<V>` (`Exclude` with no constraint to verify) removed a check at fourteen call sites.

### #13 empty-value-allowed, and the instantiation budget

After the empty-value-allowed branch landed (#13, Sept 2026): **~618k instantiations, ~1.6s check time**, against **611k, ~1.6s** immediately before it on the same machine — so putting a conditional in the Named tier's unmarked return type cost ~7k instantiations and no check time, as the existing `isFormula` trait's measurement predicted. Treat a jump past roughly **750k instantiations or 2.0s** as a design problem rather than a cost to absorb.

### #14 complete-row append

After the complete-row append landed (#14, Sept 2026): **~690k instantiations, ~1.68s check time**, against **619k, ~1.56s** immediately before it on the same machine, leaving the budget above intact. Nearly all of the ~72k is *one line* — the overlay inside `TableNamed.appendRowWithAllVals`, where `TN` is still generic, so relating `SheetDataValuesAll<TN>`'s key set back to `ColumnName<TN>` forces the filter to be evaluated at `TN`'s **constraint**: the whole Table-name union. It is the same "filter at widened Table name" cost the endpoint entry above pays, reached through a method body rather than a dispatch boundary. Measured four ways: declaring the type and naming it in the signature costs **~1.8k**; relating it at a concrete Table name costs **~12**; relating it while the Table name is generic costs the other **~71k**, and costs it either way — through the `as` or through a hand-written overlay loop on `RowNamed`; and refusing to relate it at all drops back to ~1.8k. So the type is nearly free at every call site, and one generic consumer inside the Named tier was paying for all of them. A hand-rolled filter testing only `isFormula`, skipping the redundant value-name conditional, measured ~680k — ~10k off that same generic relation, not worth a duplicated filter.

### The `appendRowWithAllVals` cast through `unknown`

**That overlay now casts through `unknown`, and the ~70k is gone** — ~689k down to ~619k, re-measured on top of #15. Nothing went unchecked; the check moved somewhere cheaper. Both bags are built from `SheetDataValues`, so the claim is true by construction rather than by luck, and the Named suite's identity assertions already pin the parameter's exact shape at real Table names for ~12 instantiations. With the cast in place, dropping the `id` exclusion still breaks three of those assertions — verified by doing it, not assumed. Profiling every other cast-shaped site in the repo (`tsc --generateTrace` plus `@typescript/analyze-trace`) turned up nothing else worth converting: the next-largest saving measured **29** instantiations, and two candidates the trace flagged saved exactly zero. This is one line, not a technique to spread. The house style's `as` rules give the condition that licenses it.

### #18 Value Config literals

After a Value Config's members became literals (#18, Sept 2026): **~656k instantiations, ~1.71s check time**, against **615k, ~1.67s** immediately before it on the same machine — so preserving the literals costs ~41k instantiations and no measurable check time. `makeValueConfigs` inferred its argument without `const`, which widened every generated member list to `string[]` and every dropdown column's value type to bare `string`; adding `const` to the type parameter restores the union the Value Config sheet actually declares. That is what makes an exhaustive switch on a dropdown column possible at all — the build-ledger endpoint routes a charge reduction on its two declared descriptions, and a third member added to the sheet is now a compile error at that switch rather than a mis-rendered line.

After a value config became either a fixed list of any primitive or a live reference (#145, Oct 2026): **~753.2k instantiations** in the app's program, against **~749.4k** before it, and **~116.4k** against **~115.7k** in the framework's, all exactly reproducible on TypeScript 5.9.2. So the kind conditional on each value config, and the wire type that follows its members, cost ~3.8k where all twelve of real-estate's lists are strings. Check times on that machine swung from 2.7s to 4.9s across identical runs, so none is recorded. The `const` guarantee survived the new union: `valueKinds/valueKinds.test.ts` fails to compile without it.

### #15 branded `DateSerial`

After the date value type became a branded `DateSerial` (#15, Sept 2026): **~689k instantiations, ~1.67s check time**, against **690k, ~1.69s** immediately before it — so the brand cost nothing, and in fact measured ~1.7k *below* its predecessor. A brand is an intersection on one value type, not a mapped type over the column unions, so it doesn't scale with the column count. (This entry first recorded the predecessor as 651k and the brand as a ~37k cost. Re-measured from clean checkouts of `ee7fc17` and `6d64991`, twice each on TypeScript 5.9.2, the before figure is 690,317 and the after 688,601, both exactly reproducible. The 651k reading matches neither commit and could not be reproduced at either.)

### #131 config types through `Register`

After the tiers took their config types from the `Register` the app augments (#131, Sept 2026): **~795.5k instantiations, ~2.31s check time**, against **795.1k, ~2.31s** immediately before it on the same machine, so reading `Configs["columnConfigs"]` in place of `typeof columnConfigs` costs ~400 instantiations and no check time. The one trap was a cycle, not a cost: the generated literals are contextually typed by `makeColumnConfigs`' constraint, so that constraint can name no `Register`-derived type (`TableNameSimple`, `ValueName`). It takes a `VN extends string` parameter instead, which keeps each `valueName` a literal, and the `valueName` check moves to `ColumnConfigsGeneric`, which `configRegister.test.ts` asserts the registered column configs satisfy and a bad `valueName` fails.

### #222 intended-state types

After the intended-state types (`PendingAlterations`, `ClaimedTraits`, #222, Nov 2026): **~130.8k instantiations** in the framework's program, against **~128.6k** at the branch's base `5a982df`, and **~786.9k** in the app's, against **~786.7k**, on TypeScript 5.9.2. So the types cost **~2.2k** in the framework, whose dev map names every tab, and **~0.2k** in the app, whose map is still empty; the ticket's cap is ~10k. A first cut that camelcased each tab's title with `SentenceToCamelCase` cost ~13.3k in a standalone program, about 340 instantiations a title, paid when the file is checked, used or not; a key-remapped mapped type cost more, not less. So the regen writes each tab's name into `idNameMap.ts`, as it does a Table's, and the types read it.

### #262 cheaper column-map types

After three changes to the column-map types (#262, Nov 2026), the app's program dropped from **789,199 to 577,646 instantiations** (−26.8%), on TypeScript 5.9.2 at base framework `54a8b21` and real-estate `4106438`. Each count reproduced exactly across three runs. Check time is the median of three runs, before and after taken one after the other in one session on the same machine:

| Program | Before | After | Check time before | Check time after |
| --- | ---: | ---: | --- | --- |
| `packages/real-estate` | 789,199 | 577,646 | 2.54s (2.28–2.70) | 2.23s (2.15–2.26) |
| `packages/framework` | 140,653 | 132,258 | 1.14s (1.11–1.42) | 1.40s (1.37–1.48) |
| `packages/framework/scripts` | 102,831 | 96,778 | 0.72s (0.72–0.77) | 0.67s (0.63–0.80) |
| `packages/framework/valueKinds` | 102,545 | 95,859 | 0.69s (0.67–0.69) | 0.62s (0.59–0.62) |
| root | 45,079 | 45,079 | 0.75s (0.72–1.16) | 0.75s (0.72–1.16) |
| `runner` | 36,721 | 36,721 | 0.81s (0.74–0.84) | 0.81s (0.74–0.84) |
| `packages/utils` | 19,273 | 19,273 | 0.21s (0.18–0.22) | 0.21s (0.18–0.22) |
| `config` | 5,193 | 5,193 | 0.09s (0.09–0.10) | 0.09s (0.09–0.10) |

The last four programs compile no framework file, so their trees are the same before and after, and one set of three runs, taken after, stands for both. Their counts match #261's baselines at the same base.

The three changes, in the order of what they save:

- **`ColumnConfigsFlat` is one mapped object** over `FlattenTwoLevels`' result, rather than the intersection of one object per Table that `FlattenTwoLevels` returns. Every full-name lookup reads it, and it saved most of the total.
- **`ColumnFullName` skips the filter when neither parameter narrows it**, and otherwise filters by key remapping (`as`) rather than indexing a mapped object by every full name.
- **`Endpoints` reads each entry from `EndpointsAll`** and drops the config-sheet keys with one template literal, a Floor tab name plus any string, rather than rebuilding each entry and testing its Table per key. The saving comes from `frameworkEndpoints`' spread into `EndpointsAll` sharing one resolved map, not from the relation: casting that spread through `unknown` saved nothing.

The framework's check time rose while its count fell, and that wasn't traced. Its count includes ~30 for a new assertion that pins each `Endpoints` entry to its own column's sheet, which nothing pinned before. A second assertion, that the entry isn't another sheet's, cost ~5k and was dropped: the first fails on its own when the entry widens. The tracing, the per-type credit and the candidates measured and dropped are in [#262's finding](https://github.com/byronbroughten/sheets-framework/issues/262#issuecomment-6752060986). A further ~26k in the app's program is available by casting `CheckboxColumnOperator`'s checkbox comparison through `unknown`; it wasn't taken, since it removes a compile-time check.

Breaking any of the three changes leaves the app's program compiling; only the framework's own type-level tests, against its dev configs, fail.

### #78 camelCase-aware `SentenceToCamelCase`

After `SentenceToCamelCase` learned the shared splitter's camelCase and acronym boundaries (#78, Oct 2026): **~112.3k instantiations, ~0.9s check time**, against **~105.4k, ~0.89s** at its parent, both exactly reproducible on TypeScript 5.9.2. With `Str.test.ts` removed from both trees the gap is **~0.9k**, so nearly all of the ~6.9k is the eleven type assertions added there; the `\s`-sized whitespace union roughly doubled the non-test share from 393. No production call site passes a literal, and `string extends S` now returns `string` for a widened name before any character walk, so the type stays on the runtime's rules rather than requiring sentence-case headers.
