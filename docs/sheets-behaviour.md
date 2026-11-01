# Sheets behaviour

An [observation index](https://github.com/byronbroughten/byro-repo/blob/master/docs/agents/prose-files.md#terms) of how Google's platforms behave (the Sheets UI, the Sheets REST API, Apps Script with `SpreadsheetApp` and triggers, and the Google auth they share), measured costs included, never the framework's own behaviour. It is grouped by subject and follows the [observation-index rules](https://github.com/byronbroughten/byro-repo/blob/master/docs/prose-style.md#observation-indexes).

Whoever settles or overturns an observation adds or updates its line in the same change as the finding.

Write-ups sourced from Google's docs live on `research/*` branches, listed by `git branch -r --list 'origin/research/*'`.

## Protections

### Lock editors

- Observed: A lock added through REST API v4 naming no editors read back with the spreadsheet's editors (two gmail users and a service account), and in the Sheets UI the owner edited its cells (Num) with no prompt and no block; Table data column on `test`; 2026-09-17; [recorded in a03c0a6a](https://github.com/byronbroughten/byro-repo/commit/a03c0a6af3e6ba86a5eafcbc229d4a90333d06af).
- Observed: A lock added through REST API v4 naming users read back with exactly those users, none of the spreadsheet's other editors added; `test`; 2026-09-17; [recorded in a03c0a6a](https://github.com/byronbroughten/byro-repo/commit/a03c0a6af3e6ba86a5eafcbc229d4a90333d06af).
- Observed: With a lock added through REST API v4 naming one gmail address on Column currency's bounded data range, the Sheets UI blocked the spreadsheet owner's edits to those cells, and an edit to the last row, which sat outside the range, went through; the address sent and the address read back were not kept; `test`; 2026-09-17; [recorded in a03c0a6a, corrected in 28fbbd79](https://github.com/byronbroughten/byro-repo/commit/28fbbd792bea56805b6edc54646280ef60fe50de).
- Observed: The Sheets UI refused to create a lock that would lock the acting user out, and REST API v4 accepted such a lock; the sheet was not recorded; 2026-09-17; [recorded in 28fbbd79](https://github.com/byronbroughten/byro-repo/commit/28fbbd792bea56805b6edc54646280ef60fe50de).

### Reading protections back

- Observed: A REST API v4 `getByDataFilter` returned the same `protectedRanges` payload as a plain get with a `protectedRanges` fields mask; `test`; 2026-09-17; [recorded in a03c0a6a](https://github.com/byronbroughten/byro-repo/commit/a03c0a6af3e6ba86a5eafcbc229d4a90333d06af).
- Observed: An open-ended warning added through REST API v4 (start row and column bounds, no `endRowIndex`) read back with no `endRowIndex`, as `G5:G`; Table data column on `test`; 2026-09-17; [recorded in a03c0a6a](https://github.com/byronbroughten/byro-repo/commit/a03c0a6af3e6ba86a5eafcbc229d4a90333d06af).
- Observed: A warning added through REST API v4 with no editors read back with an `editors` list; `test`; 2026-09-17; [recorded in a31f1523](https://github.com/byronbroughten/byro-repo/commit/a31f1523b33625f3fe9ca136ac995703f232aa71).
- Observed: After the probe's four protections were deleted, a REST API v4 get with a `protectedRanges` fields mask returned `test` with no `protectedRanges` field; 2026-09-17; [recorded in a03c0a6a](https://github.com/byronbroughten/byro-repo/commit/a03c0a6af3e6ba86a5eafcbc229d4a90333d06af).

### Rows added below a protection

- Observed: Bounded data-column protections added through REST API v4 (`B5:B16`, `E5:E16`, `I5:I16`) kept their end row through an `appendCells` that added Table row 17 and through typing below the Table in the Sheets UI, and the appended row was unprotected in the UI; Table on `test`; 2026-09-17; [recorded in a03c0a6a](https://github.com/byronbroughten/byro-repo/commit/a03c0a6af3e6ba86a5eafcbc229d4a90333d06af).
- Observed: An open-ended warning added through REST API v4 (no `endRowIndex`) reaching below the data prompted in the Sheets UI on a cell under the Table, and on a cell in a Table row added afterwards; Table on `test`; 2026-09-17; [recorded in 28fbbd79](https://github.com/byronbroughten/byro-repo/commit/28fbbd792bea56805b6edc54646280ef60fe50de).
- Observed: A protection with a concrete end row kept that end row when 1000 rows were added to the sheet; `test`; 2026-09-17; [recorded in 28fbbd79](https://github.com/byronbroughten/byro-repo/commit/28fbbd792bea56805b6edc54646280ef60fe50de).
- Observed: `$E5:E` typed into the Sheets UI's protected-range editor was saved with a concrete end row, the sheet's last row at the time; `test`; 2026-09-17; [recorded in 28fbbd79](https://github.com/byronbroughten/byro-repo/commit/28fbbd792bea56805b6edc54646280ef60fe50de).

### Warning prompts

- Observed: A warning added through REST API v4 on the ID column prompted the spreadsheet owner on an edit in the Sheets UI; Table data column on `test`; 2026-09-17; [recorded in a03c0a6a](https://github.com/byronbroughten/byro-repo/commit/a03c0a6af3e6ba86a5eafcbc229d4a90333d06af).
- Observed: A cell carrying both a lock that permitted the editing user and a warning prompted on an edit in the Sheets UI; `test`; 2026-09-17; [recorded in 28fbbd79](https://github.com/byronbroughten/byro-repo/commit/28fbbd792bea56805b6edc54646280ef60fe50de).
- Observed: A warning's `description` showed in the Sheets UI's Protected sheets and ranges sidebar and not in its prompt, which carried no custom message; `test`; 2026-09-17; [recorded in 28fbbd79](https://github.com/byronbroughten/byro-repo/commit/28fbbd792bea56805b6edc54646280ef60fe50de).
- Observed: With the probe's warnings and locks on its data columns, the Sheets UI refused a column drag for every column; `test`; 2026-09-17; [recorded in a03c0a6a](https://github.com/byronbroughten/byro-repo/commit/a03c0a6af3e6ba86a5eafcbc229d4a90333d06af).

### Whole-sheet warnings and unprotected ranges

- Observed: A whole-sheet warning prompted in the Sheets UI on a tab rename, a tab delete and a column insert; sheet and how the warning was added not recorded; 2026-09-18; [recorded in 319e5673](https://github.com/byronbroughten/byro-repo/commit/319e5673bd29b9546b1f75c8246ef42d2f9c4b15).
- Observed: An edit in the Sheets UI inside one of a whole-sheet warning's unprotected ranges did not prompt; sheet and how the warning was added not recorded; 2026-09-18; [recorded in 319e5673](https://github.com/byronbroughten/byro-repo/commit/319e5673bd29b9546b1f75c8246ef42d2f9c4b15).
- Observed: A column inserted strictly inside a whole-sheet warning's unprotected range was inside that range afterwards, and one inserted at either of its edges was outside it; sheet, how the warning was added and the inserts' surface not recorded; 2026-09-18; [recorded in 319e5673](https://github.com/byronbroughten/byro-repo/commit/319e5673bd29b9546b1f75c8246ef42d2f9c4b15).
- Observed: On a whole-sheet warning, a column insert in the Sheets UI inside a whole-grid unprotected range prompted, and a column insert strictly inside a whole-column unprotected range did not; sheet and how the warning was added not recorded; 2026-09-18; [recorded in 319e5673](https://github.com/byronbroughten/byro-repo/commit/319e5673bd29b9546b1f75c8246ef42d2f9c4b15).

### Row and column shifts

- Observed: A protection's unprotected range of `B6:D9` read back as `B5:D8` after a row above it was deleted; live spreadsheet, with the sheet, the delete's surface and whether a Table held the range not recorded; 2026-09-19; [recorded in c3dedd7a](https://github.com/byronbroughten/byro-repo/commit/c3dedd7ab940dae089feb7ab688fc482b55d850f).
- Observed: REST API v4 row and column inserts and deletes moved the sheet's protected ranges with the rows and columns; scratch tab on the dev spreadsheet; 2026-09-26; [recorded in db5ba74d](https://github.com/byronbroughten/sheets-framework/commit/db5ba74d35cb24b9cf97975a1d86a589b88c7a76).

### Deleting protections

- Observed: A REST API v4 `deleteProtectedRange` with an id no sheet held was rejected with `No protected range with id`; scratch tab on the dev spreadsheet; 2026-09-26; [recorded in 19cc66e2](https://github.com/byronbroughten/sheets-framework/commit/19cc66e2e3f7de98a3c0edb2a66a6e855e9d0f8c).
