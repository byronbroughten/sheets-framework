# Conditional format rules

Map fragment. Sibling headings live in this folder.


A sheet's rules are an ordered list and the first match wins per format property. Adds prepend at index 0; there is no caller-supplied insertion index in this version. Identity is content — ranges, condition and format — not position. A newly added rule that is already present identically over the same range queues nothing.

Ranges are the live Table data range of the structure named (sheet, column, or cell), bounded rather than open-ended. Sheets grows a rule's range when rows are inserted inside it, and a table-aware append is such an insert.

Custom formulas are stored as written for the range's top-left cell and re-evaluated relative per cell. Callers compose formula text; `anchoredA1` resolves a column name to that top-left A1 so no call site hardcodes a letter.

Removal is by exact range or by exact content. Overlap is not a match. The framework never deletes a rule whose range or content it was not asked to name, which is what keeps hand-drawn rules safe.

Every rule the payload holds is mapped, one for one, in order. A condition or format outside the modelled slice becomes an explicit unmodelable value that still occupies its index, so deletes by index cannot silently target a neighbour. Content equality against an unmodelable rule is always false; a range-scoped removal of one is allowed because the caller named the range.

Rule writes embed row and column coordinates, so they refuse to queue when that sheet's row indexes are stale. A flush that added or deleted a rule sets a third stale axis — conditional-format indexes — and a second rule-mutating flush against that sheet is refused until a refetch of its rules clears the flag.

Rules are read with a plain spreadsheet get, never `getByDataFilter`, which drops `conditionalFormats` from its reply. Google also omits an empty rule list, so a sheet missing from that read's rules has none, not "unfetched"; it omits zero-valued range fields the same way, which read back as 0.
