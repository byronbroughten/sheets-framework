# Give the model room for the states that actually occur

Design reasoning. The one-line principle lives in [`docs/design.md`](../design.md); this file holds the argument and its instances.

## The argument

A state model that can't express a real condition doesn't omit it — it *misreports* it as one of the states it does have. Before settling a model, ask which real-world conditions have nowhere to go.

## Instances

A run killed mid-flight — an Apps Script timeout, a quota kill — runs no `finally`. Under the old boolean, it displayed the *previous* run's `TRUE`: a state with no representation became a confident lie. The colour model leaves that run yellow, which says "started, never reported back" (#4, `ac7a795`). Every cell's value type includes `""`, because an untouched cell is empty rather than defaulted — a `boolean` column reads `boolean | ""`, and code that branches on it has to say what empty means instead of assuming the base type (docs/vocabulary.md, "Values").
