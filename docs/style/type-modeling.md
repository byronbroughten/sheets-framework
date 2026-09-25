# Type modeling: reasoning and examples

Disclosed from [`docs/style.md`](../style.md), "Type modeling". The rules are there, one line each; this file holds the why.

## Identity checks, never assignment

`const x: Expected = valueOfNewType` proves nothing about a mapped or conditional type: it passes against `any` and against `never`. Use the identity-based `IsExactly`/`assertType` pair from `src/testSupport/typeAssertions.ts`, and `assertNotType` to claim that two types are _not_ identical. `assertNotType` is the only way to state that a branded type like `SerialDate` is not just `number` (#15). Two corollaries, both learned the hard way:

- **A probe that needed an `any` to compile has proved nothing.** Intersecting to satisfy an indexer (`(T & Record<K, any>)[K]`) resolves to `any`, so every assertion downstream of it passes vacuously. If a type won't index without that workaround, fix the type rather than casting past it: carry the data inside the entry so the key is provably present.
- **Measure before adopting a mapped type over the config unions**, with `npx tsc --noEmit --extendedDiagnostics`. [`type-check-cost.md`](../architecture/type-check-cost.md) has the baseline and the one known cliff.

## The framework's generic abbreviations

The abbreviations are `SN` (SheetName), `VN` (ValueName), `CN` (ColumnName), `UN` (UniformRowName), `IF` (IsFormula) and `TN`. The domain-free utilities that keep bare `T`/`K`/`V`/`O` are `utils/` (`utils/Obj.ts`, for one) and `appUtils/`. Why two letters: `@byronbroughten/config`'s `docs/style/type-modeling.md`.

## The three accepted `as` idioms

External values, such as Sheets cell data, go through `Val.validate.*`/`Val.is.*`; a cast is only for data that is already runtime-safe. The accepted idioms:

- Seed a fully-typed empty accumulator up front, then fill it: `{} as SheetColumnNamesStandard<SN>`. Don't cast at the point of use.
- Use `as any` / `as unknown as X` as an escape hatch only inside low-level structural utilities (`utils/Obj.ts`, `utils/Arr.ts` and similar) that do generic structural-typing gymnastics. This is no licence to use it elsewhere; lint rejects explicit `any` everywhere but `Obj` and `Arr` and their subfolders.
- Use `as unknown as X` in ordinary code **only to buy back type-check time, and only when a test already proves the same thing more cheaply.** Both conditions are required. The cost condition: the cast must remove real, measured work. Run `npx tsc --noEmit --extendedDiagnostics` before and after; if the saving isn't in the tens of thousands of instantiations, don't cast. The proof condition: a test elsewhere must already check the exact shape the cast claims, written against one named sheet rather than a type parameter. That test is what keeps the cast from being a hole. `SheetNamed.appendRowWithAllVals` is the only place in the repo that qualifies (#14). [`type-check-cost.md`](../architecture/type-check-cost.md) has the numbers, and the profile that found no second candidate.

Test files are separately mid-migration off `as` via the `migrate-to-shoehorn` skill. That is in-progress project state, not a rule that contradicts these.

## Registries take a plain annotation

`makeStructuredConfig` infers the literal into a type parameter. That inference accepts an unknown key without error whenever at least one _valid_ key sits beside it in the same literal. A bogus key on its own does error, which is exactly what makes the hole easy to miss. That's how `occupancy_buildLedgerRunTimeLastRan` (no such column) reached the endpoint map and still passed `npm run tsc`. `export const businessEndpoints: Endpoints = { ... }` restores excess-property checking and reports the typo with a "did you mean". Keep `makeStructuredConfig` for the generated config files, where the generator supplies the keys instead of a hand-typed literal.

## Where utility types live

Custom generic utility types live in `utils/Obj.ts`, PascalCase, one transform per name: `StrictOmit`, `DistributiveOmit`, `StrictPick`, `PickStartsWith`.

`NotEmpty<V>` is the one deliberate exception to `utils/Obj.ts`: it sits in `00_Source/CellValues/cellValues.ts` beside the wire value types, because the blank it removes is the cell blank those types define, not a general structural transform (#12).
