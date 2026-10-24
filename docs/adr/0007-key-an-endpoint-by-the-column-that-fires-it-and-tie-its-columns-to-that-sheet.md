# Key an endpoint by the column that fires it, and tie its columns to that column's sheet

An endpoint's shape used to follow its column's suffix, so a runner and a selector were two kinds with fixed feedback columns and a stem their names had to rhyme on. Now an endpoint is one entry keyed by the full name of any column whose action-row checkbox fires it, holding an action and optional `timeLastRan`, `runStatus` and `selector` columns, each typed to the entry column's own sheet and to the value type it needs, so a cross-sheet column, a wrongly typed one or a key that isn't a column fails to compile rather than being checked at run time. The correlation was measured before it was chosen, at about 200 type instantiations, because it recovers a sheet from a full name by the cheap indexed lookup and never builds a full name from a sheet.

Source: byronbroughten/byro-repo#5
