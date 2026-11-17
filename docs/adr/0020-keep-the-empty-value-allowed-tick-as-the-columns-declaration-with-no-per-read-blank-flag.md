---
status: accepted, amended by sheets-framework#219, which moves the declaration from the Column Config tick to the `emptyValueAllowed` trait
---

# Keep the Empty value allowed tick as the column's declaration, with no per-read blank flag

Dropping the tick for a flag on each read, such as `value(name, "emptyAllowed")`, was considered and rejected, because it would save only the `gen:configs` run when a tick changes, a code change still needing a build and a push. Against that, the declaration buys three things a flag can't: `tsc` lists every read that must handle a blank when a tick flips, the whole-row check of blank required columns has a record for every column, and the write side lets a blank through only on a ticked column. So the tick stays, taking effect at build time from the generated configs, and `valueOrEmpty`, `valueNotEmpty` and their multi-column forms are the per-read override.

## Amended

[sheets-framework#219](https://github.com/byronbroughten/sheets-framework/issues/219) removes Column Config and its tick. The decision stands with the column's `emptyValueAllowed` trait in the generated configs as the declaration, carried forward from run to run and changed by an `alterations.ts` entry that a `sync` run lands. A flag would now save only that `sync` run, and the three things above still hold.

Source: byronbroughten/sheets-framework#149
