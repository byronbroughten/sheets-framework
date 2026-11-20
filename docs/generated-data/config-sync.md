# How sync plans and sends each alteration

Part of [generated data](../generated-data.md). `sync` plans the pending alterations in `alterations.ts` against the previous configs and the live sheet, sends the plan only with `--send`, and emits every generated file from what it read, or from a refetch after a send. No sheet holds a config: the previous configs and `alterations.ts` hold every trait the framework declares.

## The three sides of the plan

`SyncCoordinator.run()` takes API access from a Table's presence in the previous configs, changed by a `letApiAccess` alteration ([granting and revoking](./table-configs.md#granting-and-revoking-api-access)), and `emptyValueAllowed`, a column's `header`, `formula` and `deleted`, a Table's `name` and `deleted` and a tab's `deleted` from the three-way plan over pending alterations, the previous configs and the live sheet. A bare or stopped run mints no column ID and leaves a column without one out of its configs.

## Which flush sends what

Only an unstopped send writes, in these flushes in order, each alone because every later write is indexed by the layout the one before leaves ([why](../architecture/table-columns.md)):

1. The planned sends and the missing column IDs on business sheets' header rows.
2. Column deletions and inserts, then the rules replaced, re-spanned, rewritten or added, compiled against the layout they leave.
3. Column moves, then the moved Tables' rule rewrites.
4. New Tables.

A dry run on a host that withholds flushes them too, for the entered-data gate ([sync's gate](../how-it-runs/sync.md#syncs-entered-data-gate-and---confirm)).

A live ID prefix that differs from the configs stops the run before any write (`idPrefixGuard`, run over every Table, sharing the retired prefixes with the title-derived prefix).

## Deleting a column, Table or tab

A column deletion is planned as `delete the column` and sent through `TableRaw.deleteColumn`, rightmost first on each Table. A Table deletion is planned as `delete the Table` whenever the live Table exists, and sent in the first flush through `TableRaw.queueDelete`, which leaves its tab. Its plan item names, by rule handle, the conditional-format rules wholly over the Table, each range bounded and inside the header row through the last row over its columns, which go with it; each other rule touching the Table is listed with its tab, ranges and formula as partly over it, and stays on the sheet with nothing sent.

Under Check before sending, each deleted Table gets one outside-references check listing the formulas outside it that name it, each with its column's body-row range and its formula in brace form, and its plan item counts them; they will read blank cells, since Sheets rewrites a deleted Table's references to its old body rows ([observed](../sheets-behaviour.md#creating-and-updating-tables)). The formulas come from the run's live read of every API-access Table's formula columns, Tables deleted in the same run left out, so a formula a person edited since the last sync is checked as it now stands. A formula anywhere else, outside a Table or in a Table without API access, isn't checked and nothing is read for it; whenever a Table deletion is planned a notice says so.

A tab deletion is planned as `delete the tab` whenever the live tab exists, and sent in the first flush as `deleteSheet` through `SheetRaw.queueDelete`, which takes the tab's Tables with it; the entry's other traits are sent nothing. A tab deletion gets no outside-references check yet ([#323](https://github.com/byronbroughten/sheets-framework/issues/323)). A Table or column deletion on a tab the run also deletes is refused before the read, since the tab's deletion already takes the Table.

A Table a person deleted from the sheet is no conflict: an entry naming it plans nothing, the configs drop it, its prefix retires, and the report names it under Notices. A tab a person deleted is no conflict either: its Tables drop the same way, and the report names the tab under Notices before its Tables' lines.

## Removing, replacing, adding and re-spanning conditional-format rules

A Table's `conditionalFormats.remove` handle is matched against the rule handles the configs would record from the run's read, unmodelable rules included, planned as `remove the rule` and sent in the first flush as `deleteConditionalFormatRule` at the rule's fetched index, a sheet's deletes highest index first; a handle no rule has now stops the run as no such rule, listing the Table's handles, and a removal is reported under Didn't land when the refetch finds more copies of its rule than the send should leave; a handle listed twice is one removal.

A `conditionalFormats.replace` handle is matched the same way and planned as `replace the rule`; its rule spans the body rows of its `columns` (`"all"` meaning every column), its formula compiled by `BraceFormula.compileForRule` against the planned layout, and is sent as `updateConditionalFormatRule` at the rule's index in the column deletions' flush, after its deletes and inserts ([why](../architecture/conditional-format-rules.md)). A replace handle no rule has stops the run as no such rule, and one naming an unmodelable rule stops it before any send, since that rule can only be removed; a replacement is reported under Didn't land when the refetch finds no rule with the compiled rule's handle.

A `conditionalFormats.add` entry is planned as `add the rule`, compiled the same way and sent as `addConditionalFormatRule` ahead of every rule already there; an add the refetch doesn't show is reported under Didn't land ([order and check](../architecture/conditional-format-rules.md)). A rule's `bold`, `italic` or `strikethrough`, in `add` or `replace`, is refused before the read as not sendable yet. An `add` or `replace` rule's `columns` and braces may name a column the Table's `newColumns` create in the same run, by the `columnName` its header gives (`columnNameOfHeader`); the rule is compiled against the planned layout holding it and sent after the insert. Before the read, a rule naming a column neither mapped on its Table nor created there is refused as an unknown column or brace, and one naming a column name two of the Table's new headers give is refused as ambiguous.

An unstopped run plans `re-span the rule` for each rule over a Table with a brace form that ends above the Table's last body row, its plan line carrying the `ruleRespanned` effect, and stretches it to that row ([how](../architecture/conditional-format-rules.md)). A column insert before another column, a column delete or a column move settles the rules over its Table ([how](../architecture/conditional-format-rules.md#rules-after-a-column-insert-delete-or-move)): the Table's first move line, or else its first delete or insert line, carries the `rulesRewritten` effect, a deleted column's line the `rulesRemoved` effect, and each rule with no brace form a hand-drawn rule check under Check before sending.

## Renaming a column or Table

A column `header` goes as an `updateCells` write of its header cell and a Table `name` as `updateTable` with `fields: name`, so both keep dropdown styles and rewrite `Table[Header]` references.

## The refetch after a send, and what the configs drop

When a header, a name, a formula, a tab, Table or column deletion or a rule removal, replacement, add, re-span or rewrite was sent, the run refetches after the flush and emits every file from that read, so the configs follow the new header and key the Table by its new name. The configs and the ID-name map drop whatever was deleted: a column, a Table, or a tab with every Table on it, and each deleted Table's prefix retires. A deletion the refetch still shows the tab, Table or column for is reported under Didn't land, and its entry stays in `alterations.ts`. Each sent header or name that the refetch shows Sheets kept another value for is reported under Didn't land as `set to X; Sheets kept Y`: the configs take the value Sheets kept, and the entry stays in `alterations.ts`, left out of Landed. The summary counts both (`N alteration(s) landed; M didn't land`).

Every run past the refusals returns every generated file, and `alterations.ts`'s source only when an alteration landed or a rename rewrote a pending entry's names ([the rewrite](../how-it-runs/sync.md#what-sync-writes)). There is no sync endpoint: `sync` runs only on the Node host.
