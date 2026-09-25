# Model state at the granularity the concept actually has

Design reasoning. The one-line principle lives in [`docs/design.md`](../design.md); this file holds the argument and its instances.

## The argument

Store and expose a fact at the level it's *about*, not the level it happens to arrive at. The wire format's granularity is not the domain's granularity, and neither is the storage medium's.

## Instances

`isFormula` and `numberFormatType` are column-wide traits that live on the column, even though they can only be observed by sampling the top data cell — the API delivers them cell-by-cell, but they aren't cell facts (`docs/style/class-shape.md` keeps that one as a worked example, since it turns on where a member is declared). An endpoint is one concept, so it's addressed by one key — the column whose checkbox fires it is the same fact as the column that identifies it, rather than a name plus a separate registration (#5). A run's outcome is one fact, so it lives in one cell's background colour rather than a separate boolean column that could fall out of step with the timestamp beside it (#4, `ac7a795`). Checkbox-ness is a fact about a *type*, not about each column that has it, so "an untouched cell counts as unchecked" lives on the `checkbox` value name once rather than as a per-column trait repeated on every checkbox column (#12).

## Corollaries

When a container method takes an index or id that every caller already holds as its own state, the query belongs on the instance. The parameter disappearing is what turns it into a getter.
