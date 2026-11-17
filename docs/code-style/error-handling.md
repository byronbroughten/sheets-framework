# Error handling: reasoning and examples

Disclosed from [`docs/code-style.md`](../code-style.md), "Error handling & validation". The rules are there, one line each; this file holds the why.

## `Val.assert` over `!`

`!` works only at compile time: it silences the type-checker but checks nothing at runtime, so a real `null`/`undefined` crashes later, further from the actual mistake. `Val.assert(value, "label")` checks at runtime and throws immediately with a clear, labeled message (`` `${label} not found.` ``). Lint rejects `!` everywhere, including right after an explicit `if (...) throw` that already proved the value present: `Val.assert` there costs one redundant check and keeps the rule free of exceptions.

## Marked reads

The unmarked accessors carry the column's own **`emptyValueAllowed`** declaration to the read line. On a column where it is false, `value`/`valueArr` exclude `""` and throw naming the blank cell. Where it is true they hand the blank back for the caller to interpret (#13). Choosing `valueOrEmpty`/`valueArrOrEmpty` on a column that disallows blanks claims this call site has decided what a blank means, as in a check where the blank _is_ the finding. Choosing `valueNotEmpty`/`valueArrNotEmpty` on one that allows them claims this call site is stricter than its column.

Don't read through a blank-tolerant form and defer the check to a manual guard closer to where the value is used. Reading and validating in one step means a value can never be used unvalidated in between, and it validates every field in an object literal the same way.

## `try`/`catch`

`EndpointRun.run` is the only `catch` in the codebase. It exists because an endpoint's failure has to reach the sheet as a run status rather than kill the trigger. Don't generalize from it. If the question comes up elsewhere, note it as open.

