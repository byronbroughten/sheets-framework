# The chore and its dry run

What a chore is, how one is written and where it lives: [Chores](../architecture/chores.md).

```
sheets-framework chore                   # list the chores
sheets-framework chore <name>            # dry run: read the sheet, print what it would write
sheets-framework chore <name> --json     # the same, as raw request JSON
sheets-framework chore <name> --send     # apply it
sheets-framework chore <name> --send --confirm <token>   # apply it, held items included
```

**A chore's `--send` is gated on entered data as `sync`'s is** ([sync's gate](./sync.md#syncs-entered-data-gate-and---confirm)): a flush that could lose entered data is held until `--confirm` gives the token its dry run printed. Its held list goes to the package's `.probe/chore-held.json`, and a held send prints `STOPPED` and the dry run to rerun, then exits 1.

**A `--confirm` send writes one snapshot file, with the send's plan as the run printed it, and prints `undo: npm run <script> -- <snapshot>` once the send lands** (`SnapshotFolder` in `scripts/`, which `sync` is to call too). The caller names the folder with `--snapshots` and the script with `--restore-script`, the path shown from where npm was invoked; with none named the file goes under the package's `.probe/snapshots/` and the line names `restore`. Held items the snapshot can't hold are named beside the line as undone only by the spreadsheet's version history. A confirmed send first removes the folder's `snapshot-*` files whose file time is over 30 days old. A dry run and an unconfirmed send write no snapshot.

**Dry-run mode is enforced at the adapter, not at the runner and not at the chore's handle**: a dry run records each request and returns an empty response, whoever calls the flush. Why there: [ADR 0032](../adr/0032-enforce-a-chores-dry-run-at-the-node-hosts-adapter.md). Reads are not gated: a dry run fetches from the live sheet normally, and only sends are suppressed.

The preview is a rendered summary, one line per request, naming the sheet, the range and what changes (`UpdateRequestSummary`). Each capability the framework grows owes the renderer a line format — an accepted recurring cost, because a preview nobody reads converts a decision into a formality. The compiler collects it: `ModeledRequestVerb` in `GoogleSheetsAPI.ts` lists the request kinds the builders may produce, `UpdateRequestSummary` must format each one, and the entered-data gate's `FlushClassifier` must classify each one as held or free, or `tsc` fails. `--json` is the escape hatch for when a line looks wrong.