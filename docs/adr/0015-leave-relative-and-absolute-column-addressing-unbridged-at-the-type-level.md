# Leave relative and absolute column addressing unbridged at the type level

A column is addressed two ways: relatively, by the `<TN, CN>` pair the Named tier takes, and absolutely, by its `ColumnFullName`, one correlated key whose parts and filtered subsets are plain indexed lookups on the flattened column-config map. No type-level conversion was built from the pair to the full name, because a prototype of it built the whole Table-by-column cross product before it could prune, failed with `TS2590` and took type-checking from 1.19s to 6.93s; the two families meet only at runtime.

Source: byronbroughten/byro-repo#2
