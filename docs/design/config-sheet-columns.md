# A live config-sheet column must identify a row or serve a human on that sheet

Design reasoning. The one-line principle lives in [`docs/design.md`](../design.md); this file holds the argument and its instances.

## The argument

A config-sheet cell that isn't needed to find the row, and that a person doesn't fill in or read on that sheet, is a third copy waiting to disagree with the live column. Sampled generated traits do not earn a cell.

## Instances

`isFormula`, `valueName`, `hasIdColumn` and `hasNameColumn` are sampled at emit from the described column or header row, and leftover Is formula / Value title / Has ID column cells are ignored (#33). An ID prefix typed into Sheet Config is a second copy of what every column ID on that sheet already records, so the column and its uniqueness helper were removed and the prefix is sampled from the column ID row instead (#87).
