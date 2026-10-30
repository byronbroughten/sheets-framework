# A live config-sheet column must identify a row or serve a human on that sheet

Design reasoning. The one-line principle lives in [`docs/design.md`](../design.md); this file holds the argument and its instances.

## The argument

A config-sheet cell that isn't needed to find the row, and that a person doesn't fill in or read on that sheet, is a third copy waiting to disagree with the live column. Sampled generated traits do not earn a cell.

## Instances

`isFormula`, `valueName`, `hasIdColumn` and `hasNameColumn` are sampled at emit from the described column or header row, and leftover Is formula / Value title / Has ID column cells are ignored ([ADR 0011](../adr/0011-keep-only-identity-and-operator-facing-columns-on-the-config-sheets.md) states the rule a config-sheet column must meet). An ID prefix typed into Sheet Config, the tab Table Config replaced, was a second copy of what every column ID on that Table already records, so the column and its uniqueness helper were removed and the prefix is sampled from the column ID row instead ([ADR 0012](../adr/0012-assign-each-sheets-id-prefix-once-and-sample-it-from-its-column-ids.md)). [ADR 0026](../adr/0026-declare-the-table-layout-and-id-shape-in-one-framework-constant-not-on-spreadsheet-config.md) moves Spreadsheet Config's eight layout cells into a framework constant, since they identified no row and the ID and Name headers only looked like settings a person could change. Table Config, which replaced Sheet Config, is `Table ID | Table name | Sheet title | Let api access`: Table ID finds the row, Table name and Sheet title are identity a person reads there, and Let api access is the one cell a person fills in, so none of Sheet Config's leftover sampled cells came across (sheets-framework#80).
