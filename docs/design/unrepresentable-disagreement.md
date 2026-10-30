# Make disagreement structurally impossible rather than validating against it

Design reasoning. The one-line principle lives in [`docs/design.md`](../design.md); this file holds the argument and its instances.

## The argument

When two pieces of state can contradict each other, the fix is a shape where the contradiction is unrepresentable — not a check that catches it after the fact. Validation runs at one moment and reports; structure holds at every moment and can't be forgotten.

This is the most-repeated argument in the codebase, and it's what most often distinguishes the design that was kept from the one that was rejected.

## Instances

- [ADR 0011](../adr/0011-keep-only-identity-and-operator-facing-columns-on-the-config-sheets.md): sampled traits get no config-sheet cell, so there is no third copy left to disagree with the live column they're sampled from; Table Config, which replaced Sheet Config, kept the rule (sheets-framework#80).
- [ADR 0012](../adr/0012-assign-each-sheets-id-prefix-once-and-sample-it-from-its-column-ids.md): the ID prefix is sampled from the column IDs that carry it, so the "unique or empty" helper that validated a typed second copy went with the cell, and uniqueness is the generated constructor's check of sampled prefixes.
- [ADR 0007](../adr/0007-key-an-endpoint-by-the-column-that-fires-it-and-tie-its-columns-to-that-sheet.md): the entry's key carries its own sheet, so a cross-sheet or wrongly-typed column is unrepresentable rather than checked for.
- [ADR 0009](../adr/0009-carry-a-selectors-options-inside-one-selector-object.md): the selector's options live inside the selector, so an endpoint with no selector has nowhere to write one and a silent no-op is unrepresentable rather than ignored.
- An endpoint's **run state** pairs its message and its colour in one record, so no path can show one state's colour beside another's message (#4, `ac7a795`).
- The endpoint map takes a plain `: Endpoints` annotation rather than `makeStructuredConfig`, so an unknown key is a compile error ([`docs/code-style/type-modeling.md`](../code-style/type-modeling.md#registries-take-a-plain-annotation) has the hole it closes).
- Every class exposes its schema under the single name `schema`, so two accessors can't disagree about which schema a class has (`ab75c09`).
- [ADR 0010](../adr/0010-type-a-fixed-list-dropdown-column-as-its-literal-union.md): a fixed-list dropdown column's type is the union its list declares, so a member the code doesn't route is a compile error at the switch rather than a line that renders wrong.
- [ADR 0014](../adr/0014-type-a-dropdown-from-any-let-api-access-table-column-and-take-value-config-off-the-floor.md): a Name or ID column value source is deliberately left out of the literal union, since its members are rows that grow without a regen, so it is a reference typed `string` and a name with no matching row is the endpoint's to report.
- The request builders return only the kinds `ModeledRequestVerb` names, and the dry-run summary's formatter table is keyed by that union, so a request kind the framework builds but can't preview is a compile error rather than a raw-JSON line (`bd03c23`).

## Corollaries

Structure only holds where **both** sides of the contradiction live inside the type system. A Google Table's range lives in a spreadsheet a person edits by hand, so no type reaches it and no shape can make a misplaced Table unrepresentable — the per-Table placement check Identified runs after each fetch, which confirms each managed Table's `tableId`, its header row in the header zone and its column ID row, is not the second-best instrument there, it is the only one. Don't cite this principle against it ([ADR 0008](../adr/0008-check-table-placement-at-run-time-and-refuse-never-repair.md) has the check refuse a misplaced Table rather than move it back, and run per managed Table after each fetch). The stale-Table throw sits beside it for the same reason: once a flushed delete or sort has moved a Table's rows, the layout lives in the spreadsheet, so a throw on the next use of its row indexes is the only tool short of a refetch (sheets-framework#64).
Prefer narrowing a type until the bad case is unrepresentable over encoding an explanation into a fallback value. A branded-string fallback reads as friendlier but collapses back to `never` in constraint position, so it buys nothing where it matters.

This is about states that cannot be **represented**, not names that are ambiguous to a reader. Two accessors sharing a name and returning different, fully-checked types are not an instance of it — nothing is unrepresentable and the type-checker catches a mis-wiring either way, which is why, before Meta was retired (sheets-framework#63), `sheet.column(cn)` and `sheetMeta.column(cn)` were allowed to share a name (#6). The `makeStructuredConfig` case above is a hole in *checking*, which is a different failure. Citing this principle against a shared name is over-application; reach for it when a bad state can exist, not when a reader might be confused.
