import { TableOrigin } from "../01_SpreadsheetSchema/TableOrigin";

const origin = TableOrigin.expected();

// Sheet rows and columns of a Table placed where the layout expects it, for fixtures and grid reads.
export const expectedSheetLayout = {
  colIdRowIndex: origin.headSheetRowIndex("columnId"),
  groupHeading1RowIndex: origin.headSheetRowIndex("groupHeading1"),
  actionRowIndex: origin.headSheetRowIndex("action"),
  tableHeaderRowIndex: origin.headerRowIndex,
  topDataRowIndex: origin.sheetRowIndex(0),
  startTableColIndex: origin.startColIndex,
} as const;
