# Key alterations by codebase name, and have `sync` rewrite pending names when a regen renames

Every alteration in `alterations.ts` names what it changes by codebase name, as endpoint code does (Table name, column name, tab name), typed against the ID-name map, with no ID and no separate `name` field. `sync` resolves each name to an ID through the previous configs, what the agent saw at its last refresh, and rewrites a pending entry's names when a regen renames what they name. So an agent writes what it reads, with no ID lookup and no duplicated fact, and an entry held for `--confirm` or left over from a partial landing still names the thing it meant.

## The names and the rewrite

A column name is the camelCase of its header and a tab name the camelCase of its title. Display text appears only in a value that sets it (`header: "Monthly Rent"`, a new column's `header`, a tab's `title`). `sync` works on the live sheet by ID. Every regen is a `sync` run, so the rewrite sees every rename, and the run report lists each rewrite. It covers keys, `before:` and conditional-format `columns` values, rule handles changed only by a rename, and the `Table[Header]` and `${columnName}` references in pending formula text.

## What it gives up

ID keys with a `tsc`-checked `name` failed at a stale name. Name keys don't catch an agent that writes an entry from what it read before a `sync` took up a person's swap of two names: the stale name type-checks and resolves to the other column. That takes a person reusing a name another just lost, inside one agent's edit. The run report lists every name that now points at a different ID than at the last run, and the agent docs say to reread the ID-name map after a run that lists one.

## Endpoint code is not rewritten, and claims are deferred

Claims are deferred to [#286](https://github.com/byronbroughten/sheets-framework/issues/286); intended state is pending alterations alone. Endpoint and chore code are keyed by codebase name, and `sync` never writes hand-written code. A rename fails `tsc` there until retarget fixes it, with its judgment over pins and asks. When claims return, a swap that silently re-points endpoint code would make a claim compare the other column's formula and stop the run, which a rewritten claim would hide.

## Also decided

- Braces in a conditional-format formula name columns by column name (`${@moveOut}` for the row's own value, `${moveOut}` for the whole column), not header: the framework's own syntax uses codebase names, while Sheets' `Table[Header]` keeps header text.
- A regen fails when two tab titles camelCase to the same tab name, as it already does for two headers in one Table.

## Considered options

- **ID keys with a `tsc`-checked `name`** ([file shape](https://github.com/byronbroughten/sheets-framework/issues/186), [tabs](https://github.com/byronbroughten/sheets-framework/issues/209)): rejected. It costs a lookup and a second copy of every name, and its `name` goes stale on a person's rename while an entry is pending, which `sync` can fix and the agent otherwise must.
- **Keys in display text** (`"Rent"`, `"Rent Roll"`): rejected for the codebase names endpoints already use, which need no quotes where dot access applies.
- **`sync` rewriting names in endpoint and chore code**: rejected. Only a type-aware codemod could tell a column name from any other string, a rename sometimes needs retarget's ask, it would widen what a run writes into hand-written `src/`, and it would hide the claim stop that #286 may add. A codemod as retarget's first step is its own effort.

Source: byronbroughten/sheets-framework#218, byronbroughten/sheets-framework#178, byronbroughten/sheets-framework#285
