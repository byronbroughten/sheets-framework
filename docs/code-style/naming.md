# Naming

Style fragment. The one-line rules live in [`docs/code-style.md`](../code-style.md); this file holds the reasoning and the worked examples. The general naming reasoning is in `@byronbroughten/config`'s `docs/code-style/naming.md`.

One heading per naming rule, in docs/code-style.md's order: grep `^## ` for the rule you're applying and read that section.

## The profile marks a live read

**A live read with a same-named schema/config counterpart takes the bare name on the profile, never an `active` prefix** — the profile is what disambiguates the live read from the committed one. `column.profile.isFormula` (this run's live sheet data) vs. `ColumnSchema.isFormula` (the committed `columnConfigs.ts` trait) — same underlying concept, two different sources of truth. The `active` prefix that once did this job is retired (sheets-framework#63): it clashed with Apps Script's active selection.

## `_actualX` drops its prefix when a helper moves down

**When a "compute the true/live value" helper with a committed counterpart moves from a coordinating Operator down onto the column it's actually about, it lands on the profile under the bare name.** `ColumnConfigsOperator._actualValueTitle`/`_actualValidationValueName`/`_actualPrimitiveValueName` — named from the coordinator's point of view, contrasting a *live* value against the *stored config* value — became `activeValueTitle()`/`activeValidationValueTitle()` on the Meta column once they moved onto the column itself, and `column.profile.valueTitle()`/`.validationValueTitle()` once the profile replaced Meta (sheets-framework#116; #146 later generalised the latter to `.validationColumnReference()`): from that object's own point of view it's just its current state. (The innermost helper, `_actualPrimitiveValueName`, kept its old name and stayed `private` — it has no live/committed counterpart to disambiguate; see the `_`-prefix note in `@byronbroughten/config`'s `docs/code-style/naming.md` for why it's still underscored while its siblings aren't.)

## `col` or `column`

**`column` abbreviates to `col` by default — it's referenced constantly, so shortening it earns its keep (`colIndex`) — except when it's paired with an already-short suffix, where spelling it out keeps the identifier legible** (`columnId`, not `colId`: `Id` alone is too short to pair with `col` without the result reading as a cryptic blob). Whichever form fits, use it consistently within one scope: `activeColIds` next to `existingColumnIds` in the same method reads as a typo, not a style choice; it became `activeColumnIds` to match, and is now the profile's `columnIds`.

## A sheet is unmarked, a row takes `Row`

**A sheet takes the unmarked name; a row is marked with a spelled-out `Row` suffix.** `const occupancy = ss.table("occupancy")` is the sheet and `occupancyRow` is one of its rows, so the same identifier never means a sheet in one endpoint and a row in another (#22). Sheet locals outnumber row locals across the repo, which is the count "Give the common case the unmarked name" asks for before deciding which case goes unmarked, and a spelled-out `Row` reads to someone who has never opened the codebase where a coined abbreviation does not. A sheet-marking suffix was considered on the analogy of `col`, but that abbreviation marks a single column index rather than a collection, so the analogy doesn't hold.

## Google's API names stay at the wire

**Google's API names stay at the wire; everything the framework defines for itself follows the glossary** (#70). Anything that spells a Google API object, request or field keeps Google's spelling: `GoogleProtectedRange`, the queued request kinds `"addProtectedRange"` and `"deleteProtectedRange"`, the fields `protectedRangeId` and `unprotectedRanges`. Everything above that takes the glossary term, so the framework type is `EditProtection` and the sheet reads `editProtections()`. Before #70 the framework used both, and `removeEditProtection(protection: ProtectedRange)` put the two names for one thing in a single signature. The enum-string rule under the constant rule in `@byronbroughten/config`'s `docs/code-style/naming.md` is the same boundary.

## A multi-row delete is `SHOUTING_SNAKE_CASE`

**A method that deletes more than one row takes a `SHOUTING_SNAKE_CASE` name** (`TableIdentified.DELETE_ALL_DATA_ROWS` and the `TableNamed` method that delegates to it). Nothing else in the codebase is spelled that way, so the shout is the warning: a caller can't reach one by reflex. It stays shouty even once a guard makes the operation safe — the point is that the reader stops, not that the operation is unguarded.

## Collaborator names: `<Subject><Role><Tier>`

**A collaborator's role is the agent noun of its verb, a listed verb's only in its listed meaning, and its tier word comes last, like every other tier class's** (#67); the subject and the descriptive-noun fallback are `@byronbroughten/config`'s `docs/code-style/naming.md`, "A collaborator is `<Subject><Role>`". `SpreadsheetFlusherRaw` is named for `flush`, not `update`, because `update` means a local change and the flusher sends the batch; `SpreadsheetTableValidatorRaw` is named for `validate` because it throws on failure. "Handler" and "Manager" are never allowed: they fit any class, so they tell the reader nothing about which job this one does.

## The controlled verb list

This is the list's one home. **It reserves each listed verb for its meaning, and each listed meaning for its verb**: a framework method never uses a listed verb for anything else, nor a new verb for a meaning already on it. A verb the list doesn't name is free in its everyday sense, as `build` in `IdNameMapBuilderOperator` and `rewrite` in `SpreadsheetRefusalRewriterRaw` are, so the list grows only when a meaning needs reserving, not with each new job (sheets-framework#347's review list). A closed list made every new job claim another bare English word:
- `fetch` — actually hits the live Sheets API
- `integrate` — merges a fetched snapshot into local state; no API call, so it is never `fetch`
- `prep`/`gather` — queue state locally before a fetch (`prepFetchX` queues only; `gatherFetchX` queues *and* fetches)
- `update` — writes a local/queued change, not yet flushed
- `queue` — stages a write on a sheet's or row's write queue; the flusher gathers it into an operation. Internal to the queue layer, which never says `update`
- `append` — adds a new row
- `create` — builds something new on the sheet, a tab or a Table, with the IDs it needs
- `ensure` — idempotent guard: make this true, no-op if it already is
- `validate` — asserts an invariant, throws on failure
- `init` — factory setup
- `sync`/`flush` — coordinate multiple operators / send a batched write
- `discard` — drop queued changes without sending them; the counterpart to `flush`
- `classify` — sort a flush's requests into held and free
- `judge` — decide one request's cells against the run's reads
- `track` — follow cells across a flush's shifts
- **The list governs framework methods.**

