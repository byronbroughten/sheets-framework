# Keep the Empty value allowed tick as the column's declaration, with no per-read blank flag

Dropping the tick for a flag on each read, such as `value(name, "emptyAllowed")`, was considered and rejected, because it would save only the `gen:configs` run when a tick changes, a code change still needing a build and a push. Against that, the declaration buys three things a flag can't: `tsc` lists every read that must handle a blank when a tick flips, the whole-row check of blank required columns has a record for every column, and the write side lets a blank through only on a ticked column. So the tick stays, taking effect at build time from the generated configs, and `valueOrEmpty`, `valueNotEmpty` and their multi-column forms are the per-read override.

Source: byronbroughten/sheets-framework#149
