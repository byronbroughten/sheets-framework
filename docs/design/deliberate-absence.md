# Record a deliberate absence as deliberate

Design reasoning. The one-line principle lives in [`docs/design.md`](../design.md); this file holds the argument and its instances.

## The argument

The riskiest gap in an AI-assisted codebase is the one that looks like an oversight. An unexplained absence reads as a to-do and gets helpfully filled in; a *documented* absence carries its reason and survives.

## Instances

- The endpoint entry's selector is an anonymous nested object rather than a named `EndpointSelector<SN>`, and the run takes a structural copy of the entry rather than `Endpoint<SheetNameSimple>` — two references to one named generic type are compared by its measured variance, which the column filter leaves unmeasurable, so naming either shape breaks the dispatch boundary's widening in a file the editor never opened (#11; `docs/architecture/endpoint-dispatch.md`).
- There is no type-level bridge from relative `<SN, CN>` addressing to an absolute `ColumnFullName` — building one enumerates the full sheet × column cross product and takes type-checking from ~1.2s to ~7s with a `TS2590`. That absence is documented as deliberate, with the measurement attached (`docs/architecture/column-addressing.md`).
- There is no base class for the primary column alone; the column chain is deliberately not shaped like the sheet, row and cell chains (`docs/architecture/class-chains.md`).
- The Raw tier's `SheetCommonRaw` sits between the tier base class and the two concrete sheet classes rather than being folded into it, because the row and column base classes hang off that base and two of its members would be illegal overrides there (#6; `docs/architecture/class-chains.md`).
- The config-sheet floor's generated entries are never hand-edited: the seed declares the floor's structure and the generated files hold its identity, so a missing floor part is created from the seed and a drifted one overwritten, rather than anyone "fixing" an entry by hand (`docs/generated-data.md`).
- `IF` does not default to `false` on `ColumnFullName`, because that would silently shrink the union endpoint dispatch is keyed on.
- The Node host reaches Sheets and nothing else — no triggers, no Gmail, no Docs — and `ScriptApp`/`SpreadsheetApp` are left uninstalled so a chore reaching for one fails by name instead of half-working; the boundary is recorded as a boundary precisely so nobody reads the missing globals as an unfinished adapter (#25).
- Chores carry no tests, and the reason is written down next to the rule: a test for a one-off would be a second statement of the same thing, written by the same hand in the same hour, and its real check is the preview read before saying send (#25).
- Edit-protection writes do not expose `domainUsersCanEdit`, `UpdateProtectedRange`, or Table- and named-range-backed protections: a drifted protection is replaced, not edited; a Table-backed one can't be limited to one column (#50; `docs/architecture/edit-protections.md`).
- The config-sheet floor carries no edit locks: a lock naming no editors inherits the spreadsheet's editors and stops nobody, and the only lock that keeps another editor out names editors explicitly, which needs an address list this repo has nowhere to keep — so the floor is warnings-only, and keeping another editor out is out of scope rather than half-built (#49, #51; `docs/architecture/edit-protections.md`). Drift repair covers a protection's range and content but not its editors, so an editor added by hand goes unnoticed.
- Value Config is part of the floor, reversing #49, but carries no edit warning: a whole-sheet warning would prompt on every column insert, and its value lists are meant to change (#99; `docs/generated-data.md`).

## Corollaries

The same applies to commented-out code, which is why the general style doc's delete-dead-scaffolding rule (`@byronbroughten/config`'s `docs/style.md`) carves out an exception for it. Absence of an explanation is not evidence of absence of a reason — ask.
