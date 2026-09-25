# Type-check cost

Map fragment. Sibling headings live in this folder.

How to measure type-check cost, the baseline, the instantiation budget, the one known cliff, and what each typed change measured when it landed.

## How to measure

The generated config describes ~574 columns across ~28 sheets, and a lot of this codebase's typing is mapped/conditional types over those unions. That work is cheap until it isn't, so measure rather than guess: `npx tsc --noEmit --extendedDiagnostics` is always safe to run and reports instantiations and check time.

**Take both readings from clean checkouts of the two commits** — `git worktree add <dir> <commit>`, with `node_modules` symlinked in, is the quick way. Instantiation counts are deterministic for a given TypeScript version, so two runs of the same commit that disagree mean the trees differed, not that the number is noisy; a reading taken from a working tree mid-change is not comparable to one taken after the commit. Check times *are* noisy, so compare them only within one measurement run. A wrong "before" number is the easy mistake here, and it silently turns a saving into a cost.

## Baseline and the lazy mapped filter

Baseline (Sept 2026): **~286k instantiations, ~1.2s check time.** A mapped filter over the flat column map (`{ [K in ColumnFullName]: … }`) costs **~48k instantiations per distinct instantiation** — cheap, and cheaper *lazily* than precomputed: eagerly grouping every value name up front cost ~390k, while filtering per use costs ~48k each and only for the value names actually used.

## The known cliff: a widened template literal over both names

**The known cliff is a widened template literal over both names.** Building `` `${SN}${CodebaseNameDelimiter}${CN}` `` when both are unions enumerates the full ~28 × ~574 cross product before any intersection can prune it, which produces `TS2590: Expression produces a union type that is too complex to represent` and takes check time to **~7s**. Going the other way — from a full name to its parts, by indexed access on the flat map — costs nothing. That asymmetry is why there is no type-level bridge from `<SN, CN>` to a column full name.

## Measurements after each change

### #5 endpoint entry

After the endpoint entry landed (#5, Sept 2026): **~620k instantiations, ~1.54s check time**, against **583k, ~1.58s** immediately before it on the same machine — so the entry shape cost ~37k instantiations and no check time. Most of that is the two `ColumnNameFiltered` filters instantiated at the *widened* sheet name, which the dispatch boundary forces; the correlated `Endpoints` map itself measured ~200. Composing the two filters (`ColumnValueName<SN, FeedbackColumnName<SN>>`, to re-derive a feedback column's value type inside `EndpointRun`) cost a further ~43k on its own and was dropped — the entry already pins the value type, so the run resolves the column by `columnId` through the Identified tier instead.

### Absolute addressing family

After the absolute addressing family landed (Sept 2026): **~597k instantiations, ~1.6s check time**, against **286k, ~1.5s** re-measured on the same machine in the same session — so instantiations roughly doubled while check time barely moved. (Compare check times only within one measurement run: the 1.2s recorded above and the 1.5s here are the same baseline commit on different machines.) Most of the increase is the handful of distinct `ColumnFullName<VN, IF>` and `ColumnNameFiltered<SN, VN, IF>` instantiations the endpoint unions and the checkbox operator ask for.

### #11 selector object

After the selector became an object (#11, Sept 2026): **~583k instantiations, ~1.63s check time**, against **638k, ~1.72s** immediately before it on the same machine — so bundling the retain flag into the selector cost nothing and measured ~55k *fewer* instantiations. The saving wasn't chased down; what matters for the budget is that the reshape adds none.

### #12 `checkbox` value name

After the `checkbox` value name landed (#12, Sept 2026): **~610k instantiations, ~1.6s check time**, against **585k, ~1.61s** immediately before it on the same machine — so the new value name and its blank exception cost ~25k instantiations and no check time. Two things paid for themselves: correcting Raw's `OrEmpty` return types to declare the blank *reduced* instantiations, and replacing `StrictExclude<V, "">` with `NotEmpty<V>` (`Exclude` with no constraint to verify) removed a check at fourteen call sites.

### #13 empty-value-allowed, and the instantiation budget

After the empty-value-allowed branch landed (#13, Sept 2026): **~618k instantiations, ~1.6s check time**, against **611k, ~1.6s** immediately before it on the same machine — so putting a conditional in the Named tier's unmarked return type cost ~7k instantiations and no check time, as the existing `isFormula` trait's measurement predicted. Treat a jump past roughly **750k instantiations or 2.0s** as a design problem rather than a cost to absorb.

### #14 complete-row append

After the complete-row append landed (#14, Sept 2026): **~690k instantiations, ~1.68s check time**, against **619k, ~1.56s** immediately before it on the same machine, leaving the budget above intact. Nearly all of the ~72k is *one line* — the overlay inside `SheetNamed.appendRowWithAllVals`, where `SN` is still generic, so relating `SheetDataValuesAll<SN>`'s key set back to `ColumnName<SN>` forces the filter to be evaluated at `SN`'s **constraint**: the whole sheet-name union. It is the same "filter at widened sheet name" cost the endpoint entry above pays, reached through a method body rather than a dispatch boundary. Measured four ways: declaring the type and naming it in the signature costs **~1.8k**; relating it at a concrete sheet name costs **~12**; relating it while the sheet name is generic costs the other **~71k**, and costs it either way — through the `as` or through a hand-written overlay loop on `RowNamed`; and refusing to relate it at all drops back to ~1.8k. So the type is nearly free at every call site, and one generic consumer inside the Named tier was paying for all of them. A hand-rolled filter testing only `isFormula`, skipping the redundant value-name conditional, measured ~680k — ~10k off that same generic relation, not worth a duplicated filter.

### The `appendRowWithAllVals` cast through `unknown`

**That overlay now casts through `unknown`, and the ~70k is gone** — ~689k down to ~619k, re-measured on top of #15. Nothing went unchecked; the check moved somewhere cheaper. Both bags are built from `SheetDataValues`, so the claim is true by construction rather than by luck, and the Named suite's identity assertions already pin the parameter's exact shape at real sheet names for ~12 instantiations. With the cast in place, dropping the `id` exclusion still breaks three of those assertions — verified by doing it, not assumed. Profiling every other cast-shaped site in the repo (`tsc --generateTrace` plus `@typescript/analyze-trace`) turned up nothing else worth converting: the next-largest saving measured **29** instantiations, and two candidates the trace flagged saved exactly zero. This is one line, not a technique to spread. The house style's `as` rules give the condition that licenses it.

### #18 Value Config literals

After a Value Config's members became literals (#18, Sept 2026): **~656k instantiations, ~1.71s check time**, against **615k, ~1.67s** immediately before it on the same machine — so preserving the literals costs ~41k instantiations and no measurable check time. `makeValueConfigs` inferred its argument without `const`, which widened every generated member list to `string[]` and every dropdown column's value type to bare `string`; adding `const` to the type parameter restores the union the Value Config sheet actually declares. That is what makes an exhaustive switch on a dropdown column possible at all — the build-ledger endpoint routes a charge reduction on its two declared descriptions, and a third member added to the sheet is now a compile error at that switch rather than a mis-rendered line.

### #15 branded `DateSerial`

After the date value type became a branded `DateSerial` (#15, Sept 2026): **~689k instantiations, ~1.67s check time**, against **690k, ~1.69s** immediately before it — so the brand cost nothing, and in fact measured ~1.7k *below* its predecessor. A brand is an intersection on one value type, not a mapped type over the column unions, so it doesn't scale with the column count. (This entry first recorded the predecessor as 651k and the brand as a ~37k cost. Re-measured from clean checkouts of `ee7fc17` and `6d64991`, twice each on TypeScript 5.9.2, the before figure is 690,317 and the after 688,601, both exactly reproducible. The 651k reading matches neither commit and could not be reproduced at either.)

### #131 config types through `Register`

After the tiers took their config types from the `Register` the app augments (#131, Sept 2026): **~795.5k instantiations, ~2.31s check time**, against **795.1k, ~2.31s** immediately before it on the same machine, so reading `Configs["columnConfigs"]` in place of `typeof columnConfigs` costs ~400 instantiations and no check time. The one trap was a cycle, not a cost: the generated literals are contextually typed by `makeColumnConfigs`' constraint, so that constraint can name no `Register`-derived type (`SheetNameSimple`, `ValueName`). It takes a `VN extends string` parameter instead, which keeps each `valueName` a literal, and the `valueName` check moves to `ColumnConfigsGeneric`, which `configRegister.test.ts` asserts the registered column configs satisfy and a bad `valueName` fails.
