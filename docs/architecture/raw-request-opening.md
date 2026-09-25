# The raw request opening

Map fragment. Sibling headings live in this folder.


When a chore needs something the framework does not model yet, **`SpreadsheetRaw.gatherRawRequest` queues an opaque raw request onto the existing update queue**. The accepted type is the `RawSource` port's branded `OpaqueRawRequest`, which Raw can neither read nor build. A chore writes a typed Google `Request` and wraps it with `googleRawRequest` from the platform module, keeping Google's type checking; the adapter unwraps it when it maps the flush. It rides the same flush, and therefore inherits the dry run, the batching and the failure-path discard, while skipping the type layer entirely. Raw requests go out **last**, after appends, column inserts, fills, updates, find-and-replaces, deletes and sorts: a raw request is by definition outside the ordering rules the queue was designed around, so last is the safe default, and a caller needing it earlier has a specific reason they should state. A future Excel adapter would leave this unimplemented — it is Google-only.

It is **callable by anything, restrained by documentation rather than by placement** — hiding it in a chore-only module would be a fence with a gate in it, and would disguise a general hole as chore-private machinery. **Using it obliges filing an issue naming the missing framework capability.** That obligation is the only counterweight keeping the opening from becoming the reason the framework stops growing, and it is not mechanically enforced.

Raw's no-`columnId` rule is a **placement** rule, not only an access rule. A behaviour that needs a column's committed trait — `isFormula`, `valueName` — can't live at Raw even when everything else about it is index-shaped; it belongs at Identified, which resolves the id. What Raw can have is the live sampled counterpart on the Meta view (`activeIsFormula`), which costs a top-data-row fetch and can disagree with the committed config.

