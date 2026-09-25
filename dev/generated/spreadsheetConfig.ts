import { makeSpreadsheetConfig } from "../../src/01_SpreadsheetSchema/makeConfigs";

export const spreadsheetConfig = makeSpreadsheetConfig({
  idDelimiter: ":",
  idHeader: "ID",
  nameHeader: "Name",
  startTableColIndexBase0: 0,
  columnIdRowIdxBase0: 0,
  columnGroupHeadingRowIndexBase0: 1,
  actionRowIndexBase0: 2,
  tableHeaderRowIndexBase0: 3,
} as const);
