# Give the common case the unmarked name

Design reasoning. The one-line principle lives in [`docs/design.md`](../design.md); this file holds its instances.

## Instances

[ADR 0029](../adr/0029-let-a-columns-empty-value-allowed-tick-decide-what-the-unmarked-value-read-means.md) gives the validating read the unmarked `value`, since nearly every cell read is one where a blank means the endpoint cannot do its job, and leaves `valueOrEmpty` for the caller who has decided what blank means. [ADR 0030](../adr/0030-count-rowindex-from-the-tables-first-body-row-and-brand-the-sheet-row.md) keeps the name `rowIndex` for the body row nearly every caller means, so the rare sheet coordinate costs one word as the branded `SheetRowIndex`. [ADR 0009](../adr/0009-carry-a-selectors-options-inside-one-selector-object.md) makes each selector option a default-false directive, so a selection that clears after a successful run needs no declaration and keeping one costs one key. [ADR 0022](../adr/0022-give-a-declared-checkbox-column-a-checkbox-value-name-that-is-never-blank.md) takes the blank out of a declared checkbox column, so a selector is read with the plain unmarked accessor and a checked-row filter is a comparison rather than blank-handling.

## Corollaries

Where "the common case" differs per subject, let the subject declare it rather than picking one meaning for all of them ([ADR 0029](../adr/0029-let-a-columns-empty-value-allowed-tick-decide-what-the-unmarked-value-read-means.md) lets a column's own **Empty value allowed** box decide what its unmarked read means, since no call site could tell a legitimately blank column from the rest, and brings `valueNotEmpty` back as the marked word for a call site stricter than its column). That carries the principle further rather than reversing byro-repo#8's validating `value`: the unmarked word is the reflexive correct read on every column now, instead of only on one that cannot be blank.

The marked name is a record of intent, not merely a longer spelling. `valueOrEmpty` at a call site says blank was thought about, which is a fact a later reader cannot recover from the surrounding logic.

A tier that shouldn't offer the common case declines the unmarked word rather than reusing it for something else ([ADR 0029](../adr/0029-let-a-columns-empty-value-allowed-tick-decide-what-the-unmarked-value-read-means.md) leaves Raw only `valueOrEmpty` and takes `value` off Identified, which resolves columns by generated ID and can't see what a column declared). So `value` means one thing everywhere it exists and moving a call between tiers can't silently change its failure mode.
