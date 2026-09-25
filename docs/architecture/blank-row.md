# A sheet never ends a run with zero data rows

Map fragment. Sibling headings live in this folder.

The blank-row invariant, how the tiers split enforcing it, how the append path reuses the blank row, and the consumers that had to learn about it. Read it before adding a deletion path.

## The invariant

**The invariant: a delete that would take a sheet's data rows to zero clears that row in place instead.** A Sheets table recovers a new row's formulas, number formats, validation and colours from the rows it already has, so a sheet emptied completely loses all of them for every row added afterwards. The surviving row is the **blank row** — see [`CONTEXT.md`](../../CONTEXT.md) for the operator-facing word. Its formula cells still show whatever their formulas make of an empty row, so it reads as a live row rather than a gap, and the next row the app adds goes into it rather than beneath it, so it never sits stranded above the data. The rule is last-one-standing rather than positional: deleting a row while other data rows survive stays an ordinary delete. Reading the active Table is the read-side twin: if the Table's exclusive end row is not strictly past the first data row, the sheet has no data row at all and the read throws, naming the tab. A legal Table whose only data row is the blank row still has an extent of one.

## How the tiers split the rule

The tiers split the rule between enforcing it and upholding it:

- **Raw refuses.** `RowRaw.delete` throws when it would leave the sheet with no data rows. Raw can't clear a row, because clearing means resolving columns by `columnId`, which isn't Raw's job. What Raw does own is the count — `SheetRaw.dataRowCountAfterFlush` is the table's extent minus the deletes already queued for that sheet, since local row state holds only *fetched* rows — and the per-row reservation set behind `isReserved`/`reserve`/`release`.
- **Identified converts.** `RowIdentified.delete` clears instead of deleting when it would empty the sheet. `RowIdentified.clearValues` writes `""` into every configured non-formula column — the id column and checkbox columns included, though a cleared `checkbox` column then *reads* `false`, since that's what its blank means — mirrors those blanks into local state, releases the row's reservation, and never touches a formula cell. `RowIdentified.isBlank` is the emptiness test, `SheetIdentified.hasNoData` its sheet-level counterpart, and `SheetIdentified.DELETE_ALL_DATA_ROWS()` is the one way to empty a sheet: it deletes every data row below the top one and clears the top one, so which row survives a wipe is predictable rather than a consequence of loop order.
- **Named delegates.** `SheetNamed.DELETE_ALL_DATA_ROWS` and `RowNamed.clearValues` exist so business code never reaches through the Identified view for an everyday operation.

## A blank checkbox cell vs an explicit `FALSE`

**A blank checkbox cell and an explicit `FALSE` come apart at the blank test.** A declared checkbox column can never be blank at the value layer: its schema declares `blankReadsAs: false`, so the Identified accessor turns an empty cell into `false` (#12). `isBlank` does not go through that path. It asks Raw whether each fetched non-formula cell holds `""`, and Raw knows nothing of value names, so a cell holding an explicit `FALSE` is not empty and a row whose only content is an unticked box is **not blank**. A checkbox cell really can hold either: on 11 September 2026 `Add Property Expense`'s one data row held an explicit `FALSE` in a just-added checkbox column and nothing anywhere else, and counted as a row with data. Clearing writes `""` back, so the blank row a wipe leaves behind is blank in the strict sense while a row the operator inserted may not be. Whether the blank test should consult the same declaration is open, and deliberately unanswered: an endpoint that wants "nothing but unticked boxes counts as empty" currently has to say so itself. The same asymmetry sits on the append path, where `appendRowDefault` writes each column's declared default: `false` for a checkbox column, so a row appended to a sheet with one is never blank, while a `date` column's default is `""`, like a `number` column's. Reuse leans on the reservation rather than on the row still testing blank.

## The append path reuses the blank row

**The append path reuses the blank row, or it would be clutter.** When a sheet holds exactly one data row and that row is reusable — blank, and not already handed out by an earlier append this run — `SheetIdentified.appendRowDefault` updates that row instead of adding another, and either way reserves the row it returns. The reservation is what makes reuse safe on a sheet whose non-formula defaults are all empty: a second append can't collapse into the first one's row. Clearing releases it, so a run that wipes a sheet and then rebuilds it reuses the row it just cleared.

## When reuse is blocked

Two conditions block reuse and both are deliberate. A row queued for deletion is not the row that will survive the flush — indexes don't shift until then — so the append goes to the bottom as before. And a one-row sheet whose row was never fetched **throws**, naming the prefetch the caller owes: silently appending would recreate the stranded-blank-row outcome the reuse exists to prevent. Clearing a row the run had already fetched costs nothing to read back, because the write mirrors into state; clearing a row nothing fetched writes to the sheet but leaves no local state, so a later append in that same run still throws.

## `appendRowWithVals` vs `appendRowWithAllVals`

**Two appends sit on top of that path, and they differ only in what the type demands.** `SheetNamed.appendRowWithVals` takes a partial bag: every column the caller leaves out keeps the default the append already wrote, the id included, which is what config-sheet maintenance wants. `SheetNamed.appendRowWithAllVals` takes `SheetDataValuesAll<SN>` — the sheet's data-value map restricted to non-formula columns, then with `id` excluded — so every writable column is a **required key** and an unfinished bag is a compile error rather than an empty cell (#14). Naming `id` or a formula column in that literal is a compile error too: the id is minted by the defaulted append rather than supplied, and the cell layer would have thrown on the formula anyway. Both methods run the same path, `SheetIdentified.appendRowDefault` then an overlay, so blank-row reuse, reservation and id generation stay in one place; there is no complete-row append on Identified or Raw. A generated row id is a seven-character random suffix behind the sheet's id prefix, with no check against the ids already on the sheet, so an append never needs the target sheet's id column fetched. Existing `appendRowWithVals` call sites are not being migrated, because a call that is *meant* to fill only some columns shouldn't have to lie about the rest.

## What the required bag doesn't filter

**Two filters the required bag deliberately doesn't apply.** Action-control columns — a selector's checkbox, `timeLastRan`, `runStatus` — are required like any other writable column, since guessing which columns the system owns is worse than asking for them. And **Empty value allowed** neither makes a key optional nor narrows its value type: completeness is key presence, a blank stays a legal value wherever the column's value type already admits one, and any write-side meaning for the trait is left to #13, which moved the read path only.

## Consumers that know about the blank row

**Three consumers had to learn about the blank row.** `isBlank` and `clearValues` both read the configured `isFormula` trait by column id — the same source `appendRowDefault` uses for its defaults — so clearing, appending and the blank test can't disagree; a configured column the live sheet doesn't have is skipped, since there is nothing there to read or clear. The config Operators loop over `rowIndexesActiveWithData` rather than `rowIndexesActive`, because a fully pruned config sheet now ends with a blank row and their validating reads throw on a blank cell. And an endpoint declaring no selector hands its action `rowIndexesFullWithData`, so no action has to guard against the blank row itself; the run-status fill still covers the whole column, so a run against an emptied sheet stamps its status into the blank row.

## An unfetched row is never blank

A row nothing has fetched is never called blank. "Active" means present in the working view (see [the working view](./working-view.md#active-means-present-in-the-working-view)), so on the `triggerOnEdit` path — where no data row is fetched — every row counts as holding data and the selection is unchanged. That is the only honest answer available without paying for a read.
