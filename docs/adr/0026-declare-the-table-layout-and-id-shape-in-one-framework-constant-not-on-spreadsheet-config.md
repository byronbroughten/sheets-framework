# Declare the Table layout and ID shape in one framework constant, not on Spreadsheet Config

Spreadsheet Config's data row held eight cells: the ID delimiter, the ID and Name headers, and five indexes placing the Table and the rows above it. None identified a row or served a person: six were refused by the config sync if edited, and editing the ID or Name header without renaming that header on every sheet silently dropped every sheet's ID or Name column on the next sync. All eight became one framework constant in code, now `tableLayout`, base 0, with no per-app override until a second app needs a different layout. The live-read path, its overlay and the generated Spreadsheet Config file were deleted rather than kept as scaffolding for whatever reads Spreadsheet Config next, and a unit test of the constant replaced the runtime layout check.

Source: byronbroughten/sheets-framework#6
