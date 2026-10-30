# The chore and its dry run

What a chore is, how one is written and where it lives: [Chores](../architecture/chores.md).

```
sheets-framework chore                   # list the chores
sheets-framework chore <name>            # dry run: read the sheet, print what it would write
sheets-framework chore <name> --json     # the same, as raw request JSON
sheets-framework chore <name> --send     # apply it
```

**Dry-run mode is enforced at the adapter, not at the runner and not at the chore's handle**: a dry run records each request and returns an empty response, whoever calls the flush. Why there: [ADR 0032](../adr/0032-enforce-a-chores-dry-run-at-the-node-hosts-adapter.md). Reads are not gated: a dry run fetches from the live sheet normally, and only sends are suppressed.

The preview is a rendered summary, one line per request, naming the sheet, the range and what changes (`UpdateRequestSummary`). Each capability the framework grows owes the renderer a line format — an accepted recurring cost, because a preview nobody reads converts a decision into a formality. The compiler collects it: `ModeledRequestVerb` in `GoogleSheetsAPI.ts` lists the request kinds the builders may produce, and `UpdateRequestSummary` must format each one or `tsc` fails. `--json` is the escape hatch for when a line looks wrong. Because a dry run withholds every flush, a config regeneration that runs the floor then refetches still sees the live column types, so a generated-file preview can show the value-name churn that `--send` would prevent.
