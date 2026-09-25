# Naming

Style fragment. The one-line rules live in [`docs/style.md`](../style.md); this file holds the reasoning and the worked examples. The general naming reasoning is in `@byronbroughten/config`'s `docs/style/naming.md`.

One heading per naming rule, in docs/style.md's order: grep `^## ` for the rule you're applying and read that section.

## The `active` prefix marks a live read

**Prefix a getter `active` when it reads live/fetched sheet state that has a same-named counterpart sourced from schema/generated-config data** — disambiguates the live read from the committed one. `ColumnMetaRaw.activeIsFormula` (this run's live sheet data) vs. `ColumnSchema.isFormula` (the committed `columnConfigs.ts` trait) — same underlying concept, two different sources of truth. Matches the existing `active` vocabulary for "what's actually in the fetched state right now": `activeTable`, `activeColumnIds`, `activeSheetGids`, `activeRowIndexes`.

## `_actualX` becomes `activeX` when a helper moves down

**When a "compute the true/live value" helper moves from a coordinating Operator down onto the domain object it's actually about, rename it from `_actualX` to `activeX` to match that vocabulary.** `ColumnConfigOperator._actualValueTitle`/`_actualValidationValueName`/`_actualPrimitiveValueName` — named from the coordinator's point of view, contrasting a *live* value against the *stored config* value — became `ColumnMetaRaw.activeValueTitle()`/`.activeValidationValueTitle()`/`._actualPrimitiveValueName()` once they moved onto the column itself: from that object's own point of view it's just its current state, matching sibling getters like `activeHeader`/`activeIsFormula` on the same class. (The innermost helper, `_actualPrimitiveValueName`, kept its old name and stayed `private` — it has no live/committed counterpart to disambiguate, so `active` wouldn't fit; see the `_`-prefix note in `@byronbroughten/config`'s `docs/style/naming.md` for why it's still underscored while its siblings aren't.)

## `col` or `column`

**`column` abbreviates to `col` by default — it's referenced constantly, so shortening it earns its keep (`colIndex`) — except when it's paired with an already-short suffix, where spelling it out keeps the identifier legible** (`columnId`, not `colId`: `Id` alone is too short to pair with `col` without the result reading as a cryptic blob). Whichever form fits, use it consistently within one scope: `activeColIds` next to `existingColumnIds` in the same method reads as a typo, not a style choice; it became `activeColumnIds` to match.

## A sheet is unmarked, a row takes `Row`

**A sheet takes the unmarked name; a row is marked with a spelled-out `Row` suffix.** `const occupancy = ss.sheet("occupancy")` is the sheet and `occupancyRow` is one of its rows, so the same identifier never means a sheet in one endpoint and a row in another (#22). Sheet locals outnumber row locals across the repo, which is the count "Give the common case the unmarked name" asks for before deciding which case goes unmarked, and a spelled-out `Row` reads to someone who has never opened the codebase where a coined abbreviation does not. A sheet-marking suffix was considered on the analogy of `col`, but that abbreviation marks a single column index rather than a collection, so the analogy doesn't hold.

## Google's API names stay at the wire

**Google's API names stay at the wire; everything the framework defines for itself follows the glossary** (#70). Anything that spells a Google API object, request or field keeps Google's spelling: `GoogleProtectedRange`, the queued request kinds `"addProtectedRange"` and `"deleteProtectedRange"`, the fields `protectedRangeId` and `unprotectedRanges`. Everything above that takes the glossary term, so the framework type is `EditProtection` and the sheet reads `editProtections()`. Before #70 the framework used both, and `removeEditProtection(protection: ProtectedRange)` put the two names for one thing in a single signature. The enum-string rule under the constant rule in `@byronbroughten/config`'s `docs/style/naming.md` is the same boundary.

## A multi-row delete is `SHOUTING_SNAKE_CASE`

**A method that deletes more than one row takes a `SHOUTING_SNAKE_CASE` name** (`SheetIdentified.DELETE_ALL_DATA_ROWS` and the `SheetNamed` method that delegates to it). Nothing else in the codebase is spelled that way, so the shout is the warning: a caller can't reach one by reflex. It stays shouty even once a guard makes the operation safe — the point is that the reader stops, not that the operation is unguarded.

## The controlled verb vocabulary

**Method names draw from one controlled verb vocabulary**, each with a distinct meaning — don't invent a new verb for a meaning already on this list:
- `fetch` — actually hits the live Sheets API
- `integrate` — merges a fetched snapshot into local state; no API call, so it is never `fetch`
- `prep`/`gather` — queue state locally before a fetch (`prepFetchX` queues only; `gatherFetchX` queues *and* fetches)
- `update` — writes a local/queued change, not yet flushed
- `append` — adds a new row
- `ensure` — idempotent guard: make this true, no-op if it already is
- `validate` — asserts an invariant, throws on failure
- `init` — factory setup
- `sync`/`flush` — coordinate multiple operators / send a batched write
- `discard` — drop queued changes without sending them; the counterpart to `flush`
- **The list governs framework methods.** A business operator's public method takes its verb from the app's `packages/real-estate/CONTEXT.md` instead, so the code and the operator-facing vocabulary agree: the glossary says a ledger is _built_, so the method is `build`.

## Collaborator names: `<Subject><Role><Tier>`

**A collaborator is named `<Subject><Role><Tier>`, with the role the agent noun of a verb from the list above** (#67). The subject says what it works on, the role says which verb it performs, and the tier word comes last like every other tier class. `SpreadsheetFlusherRaw` is named for `flush`, not `update`, because `update` means a local change and the flusher sends the batch; `SpreadsheetTableValidatorRaw` is named for `validate` because it throws on failure. A job with no verb on the list gets a plain descriptive noun. "Handler" and "Manager" are never allowed: they fit any class, so they tell the reader nothing about which job this one does.

