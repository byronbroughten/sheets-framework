# Model state at the granularity the concept actually has

Design reasoning. The one-line principle lives in [`docs/design.md`](../design.md); this file holds the argument and its instances.

## The argument

Store and expose a fact at the level it's *about*, not the level it happens to arrive at. The wire format's granularity is not the domain's granularity, and neither is the storage medium's.

## Instances

`isFormula` and `numberFormatType` are column-wide traits that live on the column, even though they can only be observed by sampling the top data cell — the API delivers them cell-by-cell, but they aren't cell facts (`docs/code-style/class-shape.md` keeps that one as a worked example, since it turns on where a member is declared). An endpoint is one concept, so it's addressed by one key — the column whose checkbox fires it is the same fact as the column that identifies it, rather than a name plus a separate registration ([ADR 0007](../adr/0007-key-an-endpoint-by-the-column-that-fires-it-and-tie-its-columns-to-that-sheet.md) retires the stem whose names a runner's button and feedback columns had to rhyme on). [ADR 0001](../adr/0001-carry-a-runs-outcome-in-a-cell-colour-nothing-reads-back.md) holds a run's outcome, one fact, in one cell's background colour rather than in a separate boolean column that could fall out of step with the timestamp beside it. [ADR 0022](../adr/0022-give-a-declared-checkbox-column-a-checkbox-value-name-that-is-never-blank.md) states "an untouched cell counts as unchecked" once on the `checkbox` value name, since checkbox-ness is a fact about a *type* rather than a trait repeated on every checkbox column. Nearly everything a run does acts on a Table's body, so Raw and Identified hand out the Table, `ss.table(tableId)`, and the sheet is only its container, and no queued operation touches whole sheet rows or columns, so a delete, sort or column insert can't reach a neighbouring Table (sheets-framework#47; [ADR 0023](../adr/0023-carry-the-table-through-raw-and-identified-by-its-live-tableid-and-key-its-state-and-queues-by-table.md) keeps only per-sheet facts in sheet state and gives each Table its own `tableId`-keyed state and queues; [ADR 0019](../adr/0019-keep-every-row-delete-inside-its-table-and-give-a-table-no-row-insert.md) bounds a row delete to the Table it's about rather than the sheet row the API offers). [ADR 0033](../adr/0033-give-table-config-one-row-per-table-and-let-a-per-table-tick-decide-what-the-framework-touches.md) moves Let api access from the sheet to the Table, since access is a fact about a Table and an operator's own Table on a managed sheet stays theirs.

## Corollaries

A container method taking an id its only caller already holds as its own state belongs on the instance: [`docs/code-style/class-shape.md`](../code-style/class-shape.md#a-query-whose-parameter-its-caller-already-holds-belongs-on-the-instance).
