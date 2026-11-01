# Edit protections (edit warnings and edit locks)

Map fragment. Sibling headings live in this folder.

What an edit warning and an edit lock are, how the framework reads, matches and writes them, the ranges it names, and the conclusions drawn from the live probes. What Google did in those probes: [the Sheets index's Protections](../sheets-behaviour.md#protections).

## Warnings, locks and unmodelable protections

An **edit protection** (`EditProtection`) is an **edit warning**, an **edit lock**, or an unmodelable protection the framework leaves alone. The Google API type stays `ProtectedRange` at the adapter boundary, the way `ConditionalFormat` does. A warning prompts on an edit, the owner's included ([observed](../sheets-behaviour.md#warning-prompts)), and still lets the edit through. A lock takes optional `users` and `groups`. Unprotected ranges are representable only on a whole-sheet declaration.

## Lock editors and the owner lockout

A lock that names no editors inherits the spreadsheet's editors and does not stop them, so it is not a fence; a lock that names users reads back with exactly those users ([observed](../sheets-behaviour.md#lock-editors)). One such lock left the spreadsheet owner unable to edit the cells it covered ([observed](../sheets-behaviour.md#lock-editors)). **Why is unresolved.** The likely cause is address form rather than a broken guarantee: Gmail treats dotted and undotted addresses as one mailbox for delivery, but Google account identity uses the address as registered, so a lock naming `byron.broughten@gmail.com` on a file whose editor is `byronbroughten@gmail.com` may name a stranger — in which case blocking the owner is correct. The probe chores were never committed, so neither the string sent nor the string read back is recoverable; settle this before building anything on a list of editor addresses. The stronger claim once recorded here — that Apps Script's "the owner can never be removed" fails for a REST write — is withdrawn as unsupported. Separately, the GUI refused a lock that would lock the acting user out and the API accepted one ([observed](../sheets-behaviour.md#lock-editors)), so a developer can fence themselves out through this framework in a way the GUI would have prevented; there is no guard, because the framework does not know which account it is acting as.

## Reading protections

Reads use a plain spreadsheet get with a `protectedRanges` fields mask. `getByDataFilter` returned the same payload when probed ([observed](../sheets-behaviour.md#reading-protections-back)), unlike conditional formats; the adapter still uses the plain get so one request covers every sheet. What Google leaves out of that reply: [the working view](./working-view.md#presence-in-a-range-is-not-presence-in-a-payload). A range with no row or column bounds at all is a whole sheet. An open-ended column protection (start row and column bounds, no `endRowIndex`) reads back open-ended, and Google lists `editors` on a warning too ([observed](../sheets-behaviour.md#reading-protections-back)); a warning reads back with no users or groups, since it never uses them.

Every protection the payload holds is mapped, one for one, in order. A protection backed by a named range or a Table, or that sets `domainUsersCanEdit`, becomes an explicit unmodelable value that still carries its id and can still be removed by id. Google's `ProtectedRange` has no Table id; a Table-backed protection is unmodelable when it carries `namedRangeId`, which is the only Table signal on that resource.

## Matching: identity is content

Identity is content: range, kind, editors, unprotected ranges and description. An add queues nothing when a fetched or already-queued protection matches it on everything but editors and holds every editor it declares; a lock that names none comes back with extra inherited editors ([observed](../sheets-behaviour.md#lock-editors)), so an exact editor match would re-add it every run. Removal is by exact range, exact content, description or id — never by overlap — so a hand-set protection is left alone until it is named.

## Queueing writes and stale indexes

The write queue emits delete-protection then add-protection after the conditional-format adds and before the row deletes, which carry them up ([observed](../sheets-behaviour.md#row-and-column-shifts); [flush order](./flush-order.md#the-fixed-send-order-slot-by-slot)). Deletes are by id. Which protection writes refuse on stale row indexes, and the stale flag a protection-mutating flush sets: [stale row indexes](./stale-row-indexes.md). No column index goes stale, since a column insert lands only at the Table end. The add reply carries the new id; the framework does not integrate it, because the refetch is required.

## Named ranges a protection can cover

Named structures: the Table data range (`addEditWarning` / `addEditLock`), the whole sheet (`table.sheet.…WholeSheet`), a head row or the Table header row, a column's data range, an open-ended column range from a start row (`gridRangeFromRow` / `addEditWarningFromRow`), a whole column (`…WholeColumn`: the same getter at row 0), a column's header / column-ID / group-heading cell, and a single cell. Raw has a whole-row range helper. A whole-column range is a grid range, so it can never equal a whole-sheet range. An open-ended data range starts at the first data row; a whole-column range starts at the top of the sheet, so it holds the head rows too. The code assumes both cover rows added later; observed: [an open-ended warning starting at a data row prompted on a Table row added afterwards](../sheets-behaviour.md#rows-added-below-a-protection).

## Probe 2026-09-17: data-column ranges and appended rows

**Probe (2026-09-17), `test`, with follow-up checks the same day.** What Google did: [rows added below a protection](../sheets-behaviour.md#rows-added-below-a-protection), [warning prompts](../sheets-behaviour.md#warning-prompts) and [lock editors](../sheets-behaviour.md#lock-editors). The Table already had many data rows (last index 15, end 16), so the API append landed at the bottom (row 17), not in the blank at row 7. Cleanup deleted row 17 and the four probe protections. **Floor data-column range: open-ended**, because a bounded range kept its end row as rows were added below it and an open-ended one prompted on a row added later. Whether a range with no end row follows grid growth past the sheet's row count at creation is open, and not worth probing until something depends on it. The GUI's open-ended spelling and the API's are not the same thing ([observed](../sheets-behaviour.md#rows-added-below-a-protection)). A warning's description serves the sidebar ([observed](../sheets-behaviour.md#warning-prompts)) and this framework's matching, not the prompt. An earlier write-up of this probe recorded the opposite — that the open-ended warning did not cover the new Table row — and justified the same decision by the range merely *reading back* open-ended, which is a proxy for coverage rather than coverage itself.

## Probe 2026-09-18: whole-sheet warning with unprotected ranges

**Probe (2026-09-18), whole-sheet warning with unprotected ranges.** What Google did: [whole-sheet warnings and unprotected ranges](../sheets-behaviour.md#whole-sheet-warnings-and-unprotected-ranges). A column inserted at either edge of an unprotected range sits outside it until the next floor sync adds it — a gap the floor accepts. Why a whole-grid and a whole-column unprotected range disagree on a column insert is open.

## Check 2026-09-19: editable ranges shift with a deleted row

**Check (2026-09-19), a protection's editable range shifts with a row deleted above it** ([observed](../sheets-behaviour.md#row-and-column-shifts)). The config-sheet floor's carve-outs rely on it: it computes the self-describing rows' indexes from the rows as fetched, before the sync moves any, and the ranges follow the rows the sync deletes. Rows the sync appends land below the carve-outs, inside the open-ended tail. A self-describing row that does not exist yet (a fresh Table Config, before the sync appends its rows) is not carved on that run; the next sync carves it. There is no second pass.
