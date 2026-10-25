# Keep only identity and operator-facing columns on the config sheets

A live config-sheet column exists either so the host can identify the row or so a person can use it on that sheet: readable identity, declarations and helper formulas. Sampled generated traits don't earn a cell, so Column Config's Is formula and Value title and Sheet Config's Has ID column were deleted, and each regen samples `isFormula`, `valueName` and `hasIdColumn` from the described live columns, ignoring leftover cells while they remain. A cell the sync would overwrite was a third copy that invited edits that never stuck, and the operator gave up the glance at those traits on the config sheet for it.

Source: byronbroughten/byro-repo#33
