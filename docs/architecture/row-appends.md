# Row appends: a partial bag or every writable column

Map fragment. Sibling headings live in this folder. The path both appends run, blank-row reuse included: [blank row](./blank-row.md#the-append-path-reuses-the-blank-row).

## `appendRowWithVals` vs `appendRowWithAllVals`

**Two appends sit on top of that path, and they differ only in what the type demands.** `TableNamed.appendRowWithVals` takes a partial bag: every column the caller leaves out keeps the default the append already wrote, the id included, which is what config-sheet maintenance wants. `TableNamed.appendRowWithAllVals` takes `SheetDataValuesAll<TN>` — the Table's data-value map restricted to non-formula columns, then with `id` excluded — so every writable column is a **required key** and an unfinished bag is a compile error rather than an empty cell (#14). Naming `id` or a formula column in that literal is a compile error too: the id is minted by the defaulted append rather than supplied, and the cell layer would have thrown on the formula anyway. Both methods run the same path, `TableIdentified.appendRowDefault` then an overlay, so blank-row reuse, reservation and id generation stay in one place; there is no complete-row append on Identified or Raw. A generated row id is a seven-character random suffix behind the sheet's id prefix, with no check against the ids already on the sheet, so an append never needs the target sheet's id column fetched. Existing `appendRowWithVals` call sites are not being migrated, because a call that is *meant* to fill only some columns shouldn't have to lie about the rest.

## What the required bag doesn't filter

**Two filters the required bag deliberately doesn't apply.** Action-control columns — a selector's checkbox, `timeLastRan`, `runStatus` — are required like any other writable column, since guessing which columns the system owns is worse than asking for them. And **Empty value allowed** neither makes a key optional nor narrows its value type: completeness is key presence, a blank stays a legal value wherever the column's value type already admits one, and any write-side meaning for the trait is left to #13, which moved the read path only.
