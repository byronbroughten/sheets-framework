# Conditional format rules

Map fragment. Sibling headings live in this folder.

## Order, ranges and formulas

A sheet's rules are an ordered list and the first match wins per format property. Adds prepend at index 0; there is no caller-supplied insertion index in this version. A rule is identified by its [rule handle](#unmodelable-rules-reads-and-rule-handles), its readable content, not its position: an `alterations.ts` `remove` or `replace` names a rule by handle, sync matches the handle on its read, and a `replace` rewrites the rule at the index it found ([sync's alterations](#syncs-removals-replacements-and-adds)). The imperative API that chores call matches by content, ranges, condition and format, and an add already present identically over the same range queues nothing.

Ranges are the live Table data range of the structure named (sheet, column, or cell), bounded rather than open-ended. Growth's `PASTE_FORMAT` copy grows a rule's range ending at the Table's last row over the new rows, which its insert alone leaves out ([observed](../sheets-behaviour.md#copying-format-and-validation-onto-new-rows), [insert alone](../sheets-behaviour.md#growing-a-table-with-an-insert-and-a-widen)).

Custom formulas are stored as written for the range's top-left cell and re-evaluated relative per cell. Callers compose formula text; `anchoredA1` resolves a column name to that top-left A1 so no call site hardcodes a letter.

## Sync's removals, replacements and adds

Removal is by exact range, by exact content, or by rule handle, which sync turns into the rule's index on the read it matched the handle on. Overlap is not a match. The framework never deletes a rule whose range, content or handle it was not asked to name, which is what keeps hand-drawn rules safe.

Replacement is by rule handle too: sync rewrites the rule at its index with `updateConditionalFormatRule`, so its place in the order holds. The index is the one on the read the handle matched, less the removals of lower-indexed rules on the same sheet that the flush before it sent; the replacing rule is compiled against the Table's planned layout and sent after the run's column deletes and inserts, so Google doesn't shift it again. An unmodelable rule can't be replaced, since sync can't write what it held; it can only be removed.

Sync's adds are compiled the same way and queued on the sheet in the replacements' flush, each at index 0 and the list last first, so the list's first rule ends first. A batch sends its adds before its updates, so each replacement's index also counts the run's adds on its sheet. An add equal to a rule already there gives that rule's handle a second copy, so an add is checked on the refetch by copies, as a removal is: named on the planned layout, it expects the read's copies less those the run removes or replaces away, plus the replacements and adds that compile to it.

## Re-spanning a rule that ends short of its Table

A rule over a Table whose ranges end above the Table's last body row, as rows a person typed below the Table leave it, is re-spanned the way a replacement is sent: its record turned back into a rule, compiled against the planned layout so its ranges and whole-column references end on the last body row, and sent with `updateConditionalFormatRule` at its index in the replacements' flush, indexed the same way. Each one is planned from the live read on every unstopped run, rather than from an alteration, so it goes in the run's batch or is its only send; a rule with no brace form, one a pending `replace` or `remove` names, one a column insert, append, delete or move rewrites, one on a Table whose columns shift this run, which the next run's read re-spans, and every rule on a tab where the run deletes a Table, a column or the tab are left alone. A whole-column brace reads back whatever its end row, so a re-span keeps its handle, and one the refetch still finds short is reported under Didn't land.

## Rules after a column insert, delete or move

A column insert before another column, a column delete or a column move shifts the cells under a Table's rules, and Sheets leaves an `$`-anchored reference, or one to a deleted column, as written ([observed](../sheets-behaviour.md#column-inserts-bounded-to-a-table)). So each rule over that Table is settled from the live read, one a pending `replace` or `remove` names left to it. A rule wholly over deleted columns Sheets drops with them, so nothing is sent for it, and each later index on its sheet counts it. A brace rule whose `columns` or formula names a deleted column is removed in the first flush, before the delete reaches it. Every other brace rule is rewritten the way a replacement is sent, compiled from its record against the planned layout, which also stretches one ending short of the Table, so it isn't re-spanned. A rule with no brace form can't be rewritten, so it is listed under Check before sending as a hand-drawn rule. An append at a Table's end moves no rule, so on a Table no insert, delete or move reshapes it rewrites only each rule whose `columns` is `all`, in the first flush, widening it over the appended columns; it lists no hand-drawn rule. A moved Table's rewrites are compiled against the layout its moves leave and sent at the end of the moves' batch; when the Table also has a delete or insert, each is first sent in the first flush on the layout that leaves, since a delete leaves a whole-Table rule one column past the Table, where the move's band insert can't carry it.

## Unmodelable rules, reads and rule handles

Every rule the payload holds is mapped, one for one, in order. A condition or format outside the modelled slice becomes an explicit unmodelable value that still occupies its index, so deletes by index cannot silently target a neighbour, and that keeps its raw content: the JSON of the Google rule without its ranges. Content equality against an unmodelable rule is always false; a range-scoped removal of one is allowed because the caller named the range.

Rule writes embed row and column coordinates. Their refusal on stale row indexes, and the stale flag a rule-mutating flush sets: [stale row indexes](./stale-row-indexes.md).

Rules are read with a plain spreadsheet get, never `getByDataFilter`, which drops `conditionalFormats` from its reply ([observed](../sheets-behaviour.md#reading-conditional-formats)). `sync` reads every sheet's rules on the sheet-properties get that opens it, by `includeConditionalFormats`, so they cost it no round trip; the rules are as of that get, before the sync's flush. What Google leaves out of that reply: [the working view](./working-view.md#presence-in-a-range-is-not-presence-in-a-payload).

The regen writes each rule as a record in `conditionalFormatConfigs` ([the file](../generated-data/conditional-format-configs.md)), keyed by its **rule handle**: readable content, not a hash and not an index, so a rule keeps its key when another moves. Each rule is judged against its Table's end row as the state holds it after the sync's flush.
