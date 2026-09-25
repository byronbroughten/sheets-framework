# Make disagreement structurally impossible rather than validating against it

Design reasoning. The one-line principle lives in [`docs/design.md`](../design.md); this file holds the argument and its instances.

## The argument

When two pieces of state can contradict each other, the fix is a shape where the contradiction is unrepresentable — not a check that catches it after the fact. Validation runs at one moment and reports; structure holds at every moment and can't be forgotten.

This is the most-repeated argument in the codebase, and it's what most often distinguishes the design that was kept from the one that was rejected.

## Instances

- A sampled generated trait that also occupies a config-sheet cell is a third copy that can disagree with the live column — Column Config and Sheet Config keep only identity and operator-facing columns, and emit samples `isFormula` / `valueName` / `hasIdColumn` / `hasNameColumn` / `idPrefix` at regen (#33, #87).
- The "ID prefix is unique or empty" helper validated a Sheet Config cell that the new design removes, so uniqueness is the generated constructor's check of sampled prefixes rather than a formula watching a second copy (#87).
- An endpoint entry is keyed by a **column full name that carries its own sheet**, so `SheetNameOf<FN>` narrows every column it may declare to that sheet's, each filtered to the value type it needs — a cross-sheet or wrongly-typed column is unrepresentable rather than checked for (#5).
- Bundling `retainSelection` and `requireOneRow` inside the selector they modify is the same move one level down: an endpoint with no selector has nowhere to write either, so a silent no-op is unrepresentable rather than ignored (#11, #18).
- An endpoint's **run state** pairs its message and its colour in one record, so no path can show one state's colour beside another's message (#4, `ac7a795`).
- The endpoint map takes a plain `: Endpoints` annotation rather than `makeStructuredConfig`, which infers the literal and silently accepts an unknown key whenever a valid key sits beside it — that hole is how a nonexistent column reached the endpoint map and still type-checked (the framework's docs/style.md, "Type modeling").
- Every class exposes its schema under the single name `schema`, so two accessors can't disagree about which schema a class has (`ab75c09`).
- A dropdown column's value type is the **literal union its Value Config declares**, so code that routes on one routes exhaustively and a member added to the sheet is a compile error at the switch rather than a line that renders wrong (#18) — which is only true because `makeValueConfigs` takes a `const` type parameter; dropping it widens every member to `string` and the exhaustiveness evaporates with no error anywhere.
- The request builders return only the kinds `ModeledRequestVerb` names, and the dry-run summary's formatter table is keyed by that union, so a request kind the framework builds but can't preview is a compile error rather than a raw-JSON line (`bd03c23`).

## Corollaries

Structure only holds where **both** sides of the contradiction live inside the type system. A Google Table's range lives in a spreadsheet a person edits by hand, so no type reaches it and no shape can make a misplaced Table unrepresentable — the runtime check in `SpreadsheetRaw`'s post-fetch step is not the second-best instrument there, it is the only one. Don't cite this principle against it (#9).

Prefer narrowing a type until the bad case is unrepresentable over encoding an explanation into a fallback value. A branded-string fallback reads as friendlier but collapses back to `never` in constraint position, so it buys nothing where it matters.

This is about states that cannot be **represented**, not names that are ambiguous to a reader. Two accessors sharing a name and returning different, fully-checked types are not an instance of it — nothing is unrepresentable and the type-checker catches a mis-wiring either way, which is why `sheet.column(cn)` and `sheetMeta.column(cn)` are allowed to share a name (#6). The `makeStructuredConfig` case above is a hole in *checking*, which is a different failure. Citing this principle against a shared name is over-application; reach for it when a bad state can exist, not when a reader might be confused.
