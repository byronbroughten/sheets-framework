# Error handling: reasoning and examples

Disclosed from [`docs/style.md`](../style.md), "Error handling & validation". The rules are there, one line each; this file holds the why.

## `Val.assert` over `!`

`!` works only at compile time: it silences the type-checker but checks nothing at runtime, so a real `null`/`undefined` crashes later, further from the actual mistake. `Val.assert(value, "label")` checks at runtime and throws immediately with a clear, labeled message (`` `${label} not found.` ``). Lint rejects `!` everywhere, including right after an explicit `if (...) throw` that already proved the value present: `Val.assert` there costs one redundant check and keeps the rule free of exceptions.

## Marked reads

The unmarked accessors carry the column's own **Empty value allowed** declaration to the read line. On an unticked column, `value`/`valueArr` exclude `""` and throw naming the blank cell. On a ticked one they hand the blank back for the caller to interpret (#13). Choosing `valueOrEmpty`/`valueArrOrEmpty` on an unticked column claims this call site has decided what a blank means, as in a drift check where the blank _is_ the drift. Choosing `valueNotEmpty`/`valueArrNotEmpty` on a ticked one claims this call site is stricter than its column.

Don't read through a blank-tolerant form and defer the check to a manual guard closer to where the value is used. Reading and validating in one step means a value can never be used unvalidated in between, and it validates every field in an object literal the same way. `ColumnConfigOperator.newColumnConfigs()` reads `columnId`/`sheetGid`/`header`/`emptyValueAllowed` all via `col.x.value(rowIndex)` for exactly that reason. The drift comparisons a few methods up read `valueOrEmpty`, because a blank config cell is what they exist to catch. Sampled `isFormula` / `valueName` come from the described live column, not from Column Config cells.

## `try`/`catch`

`EndpointRun.run` is the only `catch` in the codebase. It exists because an endpoint's failure has to reach the sheet as a run status rather than kill the trigger. Don't generalize from it. If the question comes up elsewhere, note it as open.

