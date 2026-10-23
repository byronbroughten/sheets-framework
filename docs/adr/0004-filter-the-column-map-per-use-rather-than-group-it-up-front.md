# Filter the column map per use rather than group it up front

A mapped filter over the flat column map costs about 48k type instantiations per distinct use, while grouping every value name up front cost about 390k. So column types are filtered lazily, paying only for the value names actually used, measured in `docs/architecture/type-check-cost.md`.

Source: commit byronbroughten/byro-repo@ecd32996
