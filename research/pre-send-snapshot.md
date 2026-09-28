# Pre-send snapshot and restore for a Sheets harness

Research for byronbroughten/sheets-framework#27 (map #24). Question: how can a harness snapshot a spreadsheet before it sends an arbitrary `spreadsheets.batchUpdate`, and restore it afterwards?

Researched 2026-09-27 against Google's own docs. Claims tagged **[doc]** are quoted or paraphrased from the cited page. Claims tagged **[unverified]** are my inference and not stated by Google. Each one should get a one-off check on the `Sheets Framework Dev` spreadsheet before the harness depends on it.

## Answer in brief

- **Drive revisions are no use for a programmatic snapshot or restore of a native Sheet.** `keepForever` applies only to binary files. The API has no restore method. Only the UI's "Restore this version" swaps content back in place.
- **Drive `files.copy` is the cheapest complete snapshot**: one call, 50 Drive quota units, and every in-file object comes along. Its weakness is the restore: the copy has a new spreadsheet ID, so you can't swap it in. You have to bring tabs back into the live file.
- **A `spreadsheets.get` dump of only the sheets the batch touches** can be restored in place, so sheet IDs and inbound references stay intact. You replay it as `batchUpdate` requests. It needs the most code, and some IDs (protected ranges) can't be recreated.
- **Recommendation**: take a Drive copy before every send, and also record the head revision ID. When the batch only edits cells or formats, stays inside existing sheets and touches no more than about 50k cells, also dump those sheets to local JSON and restore by replaying it. Use the tab copy-back restore for anything structural or large.

## Mechanisms

### 1. Drive `files.copy` (whole-file copy)

**How the snapshot is taken.** Call `POST drive/v3/files/{spreadsheetId}/copy` with body `{name: "<title> snapshot <iso-time> <run-id>", parents: [<snapshot folder>]}`. **[doc]** It "Creates a copy of a file and applies any requested updates with patch semantics". The optional `copyComments` flag copies "the open (unresolved) comments" ([files.copy](https://developers.google.com/workspace/drive/api/reference/rest/v3/files/copy)).

**Cost, speed and quota.**
- **[doc]** Quota: an edit-class call is 50 units against 1,000,000 per minute per project and 325,000 per minute per user ([Drive limits](https://developers.google.com/workspace/drive/api/guides/limits)). Those figures are for projects created on or after 2026-05-01. Older projects keep their previous quotas. Snapshot quota is effectively free.
- **[doc]** Copies count toward the 750 GB-per-day upload limit, and the maximum file you can copy is 750 GB. A Sheet caps at 20 million cells or 100 MB ([Drive file limits](https://support.google.com/drive/answer/37603)), so neither limit binds.
- **[unverified]** Latency is typically a few seconds and grows with file size. It doesn't touch the Sheets API's 60-writes-per-minute-per-user quota.
- **Storage**: native Sheets historically don't count against Drive storage. Snapshots still pile up in a folder and need pruning, for example keeping the last N.

**What survives in the copy.**
- **[unverified]** Everything stored inside the spreadsheet: values, formulas, formats, merges, data validation, conditional formats, named ranges, protections, banding, filters and filter views, Tables, charts and developer metadata. Sheet IDs (`gid`) are most likely preserved. Table IDs, named range IDs and protected range IDs are probably preserved, but check that.
- **[doc]** Comments are copied only if `copyComments=true`, and then only the open ones.
- **[doc]** A container-bound Apps Script project is copied as a new script project: "If they make a copy of the container file … can see and run a copy of the script" ([bound scripts](https://developers.google.com/apps-script/guides/bound)).
- **[unverified]** Installable triggers are not copied. Google doesn't document this, but it is widely reported ([community thread](https://support.google.com/docs/thread/238953334)).
- **What does not survive**: sharing, beyond what the parent folder grants; version history; the spreadsheet ID; and anything outside the file that points at it by ID, such as `sheets.config.json` targets, clasp's `scriptId`, IMPORTRANGE consumers, links and bookmarks.

**How a restore works.** The live spreadsheet ID must stay the same, because configs, clasp and links key on it. So you can't "swap in the copy". The restore is **tab copy-back**:

1. `spreadsheets.sheets.copyTo` from the snapshot's sheet into the live spreadsheet. **[doc]** It "Copies a single sheet from a spreadsheet to another spreadsheet. Returns the properties of the newly created sheet." The only input is `destinationSpreadsheetId` ([copyTo](https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets.sheets/copyTo)). The new tab gets a new ID and a "Copy of …" title, and you can't pick either.
2. In one `batchUpdate`, delete the damaged tab. Then `duplicateSheet` the copied-in tab with `newSheetId` = the original sheet ID, `newSheetName` = the original title and `insertSheetIndex` = the original index. **[doc]** `DuplicateSheetRequest` lets you set `newSheetId` and `newSheetName` ([request types](https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets/request)). Finally delete the temporary "Copy of" tab.
3. Re-add the spreadsheet-level objects that lived on that tab, such as named ranges, from the snapshot's `spreadsheets.get` metadata.

Caveats of tab copy-back, all **[unverified]** and worth one dev-sheet test:
- Formulas on other tabs that referenced the deleted tab. Sheets usually shows "Unresolved sheet name" and re-resolves once a tab with that name returns, but test it.
- Named ranges pointing into the deleted tab may break or be dropped.
- Table IDs and protected-range IDs on the re-created tab will be new.
- Charts on other tabs that sourced data from the deleted tab may lose their source.

**The alternative in-place restore**: after step 1, `copyPaste` the copied-in tab onto the original with `PASTE_NORMAL`, which carries "values, formulas, formats, and merges" **[doc]**, and `PASTE_CONDITIONAL_FORMATTING`. Then delete the temporary tab. This keeps the original sheet ID and every inbound reference. It does not restore Tables, protections, banding, filters, charts, named ranges, or sheet properties such as grid size and frozen rows and columns. Those come from the JSON metadata (mechanism 4). `copyPaste` works only within one spreadsheet, which is why step 1 comes first.

A whole-file restore (all tabs) is the same procedure for every tab. For a real disaster, a human can also open the copy and use it directly, but that repoints everything at a new ID.

### 2. Drive revisions (version history)

- **[doc]** `keepForever` "is only applicable to files with binary content in Google Drive" ([revisions resource](https://developers.google.com/workspace/drive/api/reference/rest/v3/revisions); [files.copy](https://developers.google.com/workspace/drive/api/reference/rest/v3/files/copy)). You can't pin a native Sheet's revision through the API.
- **[doc]** The revisions resource has only `get`, `list`, `update` and `delete`. There is **no restore method**.
- **[doc]** "The list of revisions returned … might be incomplete for files with a large revision history" ([manage revisions](https://developers.google.com/workspace/drive/api/guides/manage-revisions)). Sheets revisions get merged over time, so a revision ID you recorded may later stop existing on its own.
- **[doc]** You can't download a Docs Editors revision's raw content. Each revision has `exportLinks` for other formats, such as xlsx. Restoring from an export means importing a new file with a new ID, and the xlsx round-trip is lossy for Sheets-only features. This is **[unverified]** for Tables, Apps Script and named functions.
- **[doc]** The UI can "Restore this version" in place, keeping the ID and history. The UI also allows "up to 15 named versions per spreadsheet" ([Docs Editors help](https://support.google.com/docs/answer/190843)). Naming a version has no API.
- **Use**: before a send, record the head revision ID and its `modifiedTime` from `revisions.list` (a 100-unit list call). It's a free breadcrumb that tells a human which version to restore in the UI. It is not a mechanism.

### 3. `sheets.copyTo` as the snapshot (copy tabs into a backup spreadsheet)

**How the snapshot is taken.** Keep one backup spreadsheet per target. Before a send, `copyTo` each sheet the batch touches into it: one Sheets write per tab.

**Cost, speed and quota.**
- **[doc]** Each copy counts against Sheets' 60 writes per minute per user and 300 per minute per project ([Sheets limits](https://developers.google.com/workspace/sheets/api/limits)).
- Fast for a few tabs.
- Worse than `files.copy` once the batch touches more than a couple of tabs.

**What survives.** **[unverified]** In the tab's grid: values, formats, validation, conditional formats, merges and probably protections, Tables and charts. Sheet IDs and table IDs change. Named ranges and anything else at spreadsheet level don't come along. Cross-sheet formulas in the copied tab now point at tabs that don't exist in the backup spreadsheet, so they show errors, but their text is preserved.

**How a restore works.** The same as mechanism 1's steps 1–3, except it copies back from the backup spreadsheet.

**Verdict.** This is strictly dominated by `files.copy`, which is cheaper on the scarce quota, takes everything at once and keeps cross-sheet formulas resolvable in the snapshot. Its only edge is keeping all snapshots in one file. A variant is a hidden `duplicateSheet` backup tab inside the live file: no second file and free `copyPaste` restore. It pollutes the live spreadsheet, though, and `gen:configs` would see the extra tab. Avoid it.

### 4. `spreadsheets.get` grid-data dump (local JSON)

**How the snapshot is taken.** Make one read: `spreadsheets.get` with `ranges=<each touched sheet>` and a `fields` mask. The mask covers:
- spreadsheet `namedRanges` and `developerMetadata`
- per sheet: `properties`, `merges`, `conditionalFormats`, `protectedRanges`, `bandedRanges`, `basicFilter`, `filterViews`, `tables`, `charts`, `developerMetadata` and `rowMetadata`/`columnMetadata` (from `data`)
- per cell: `userEnteredValue`, `userEnteredFormat`, `dataValidation`, `note`, `textFormatRuns` and `pivotTable`

**[doc]** "includeGridData … is ignored if a field mask was set". For large spreadsheets, "retrieve only the specific spreadsheet fields that you want" ([spreadsheets.get](https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets/get)). Write the response to disk under the run's directory, next to the preview.

**Cost, speed and quota.**
- One read against 60 per minute per user **[doc]**.
- **[doc]** A request running longer than 180 seconds times out. Google recommends payloads of 2 MB or less ([Sheets limits](https://developers.google.com/workspace/sheets/api/limits)).
- **[unverified]** Full-cell JSON runs about 100–300 bytes per non-empty cell. So around 50k cells gives tens of MB to read, but only a handful of 2 MB `updateCells` chunks to write back for the cells that changed.

**What survives.** Whatever the mask captures, identical to the live state. It's the most transparent format, and the agent can diff it against the post-send state.
- **[doc]** Sheet, banded-range, named-range, filter-view and chart IDs can be set on re-creation (`AddSheetRequest` and `DuplicateSheetRequest` take IDs; `bandedRangeId` and `filterViewId` are writable).
- **[doc]** `protectedRangeId` "is read-only", so re-created protections get new IDs.
- **[unverified]** Tables: Google's own `addTable` example passes `"tableId": "123"` ([tables guide](https://developers.google.com/workspace/sheets/api/guides/tables)), which suggests table IDs can be set, but it isn't stated.
- **Doesn't cover**: comments (outside the grid), Apps Script, anything outside the touched sheets, and computed values of volatile formulas (irrelevant, since they recompute).

**How a restore works (replay, in place).** Generate one `batchUpdate`, possibly chunked:
1. For each touched sheet, `updateSheetProperties` back to the snapshot: grid size, frozen rows and columns, title, index, hidden and tab color. Then `insertDimension`/`deleteDimension` so the grid matches.
2. `updateCells` over the snapshot's range with `fields: "userEnteredValue,userEnteredFormat,dataValidation,note,textFormatRuns"`. **[doc]** With `range`, it "clears fields matching those set in `fields` if data doesn't cover the entire range", so cells that were empty before come back empty. Then `unmergeCells` followed by `mergeCells` from `merges`.
3. Delete-then-re-add the sheet-level objects: conditional formats, banding, filter and filter views, protections, Tables and charts, with the snapshot's IDs wherever settable. Also named ranges on the touched sheets.
4. A sheet the batch deleted: `addSheet` with the original `sheetId`, then steps 1–3. A sheet the batch added: `deleteSheet`.

A cheaper restore for the usual mistake: diff the post-send dump against the snapshot and replay only the changed cells and objects.

## Recommendation

**Default: `files.copy` on every send, plus the head revision ID, plus a scoped JSON dump when the batch qualifies.**

| The batch… | Snapshot | Restore |
| --- | --- | --- |
| touches ≤ ~50k cells in existing sheets, with values, formats, validation, notes or conditional formats only | `files.copy` + JSON dump of touched sheets | replay the JSON in place (IDs and inbound refs kept) |
| adds, deletes or reorders sheets or dimensions, touches Tables, protections, named ranges or charts, or touches > ~50k cells | `files.copy` | tab copy-back from the copy (mechanism 1), then re-add named ranges from the copy's metadata |
| went badly wrong across the whole file | (already have both) | a human uses the UI "Restore this version" at the recorded revision, which is the only in-place whole-file restore that keeps the ID and history |

Why `files.copy` is always taken:
- It is one call and 50 units on a quota the harness otherwise barely uses.
- It captures things the JSON mask misses: comments, Apps Script, other tabs.
- A human can open it and look.

The JSON dump is extra because it is what makes the common restore safe and in place. The harness already knows which sheets a batch touches, since every request carries `sheetId`, `range` or `tableId`, so it can scope the dump without guessing.

How to set the threshold: the 50k-cell cutoff is a guess. It keeps the dump well under the 180 s timeout and the restore to a few 2 MB chunks. Tune it after timing a dump of the largest real-estate tab.

Housekeeping: keep snapshots in one Drive folder per target, named `<title> · <iso-time> · <run-id>`. Prune to the last N, or to the last N days.

## Open checks (dev sheet only)

1. Does `files.copy` preserve sheet IDs, table IDs and named-range IDs?
2. After you delete a tab and re-create it with the same ID and title, do formulas on other tabs and named ranges re-resolve?
3. Can `addTable` set `tableId`?
4. How long does a full-mask dump take on the largest tab, and how big is it?
