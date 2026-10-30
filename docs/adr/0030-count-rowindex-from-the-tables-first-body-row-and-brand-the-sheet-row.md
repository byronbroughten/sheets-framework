# Count `rowIndex` from the Table's first body row, and brand the sheet row

Row indexes count from the Table everywhere except at the Google boundary: `rowIndex` 0 is the Table's first body row through Named, Identified, Raw and the write queue, and it keeps its name, so `selectedRowIndexes`, `.row(i)` and `RowReports` didn't change for app code. The rows above the header are reached by role, not by negative index, since they are a fixed set nobody loops over. A body index becomes a sheet row only when requests are gathered, against the Table's live header, at the edit trigger's entry and in operator messages, and a sheet row is a branded `SheetRowIndex`. The body counts from 0 rather than the header, because developers handle body rows almost exclusively and a body index then matches array positions in a body read and Excel JS's body range.

Source: byronbroughten/sheets-framework#55
