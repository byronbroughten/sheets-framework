// No imports: makeConfigs.ts imports this, and generated/ imports makeConfigs.ts.

export const spreadsheetConfigTextHeaders = {
  idDelimiter: "ID delimiter",
  idHeader: "ID header",
  nameHeader: "Name header",
} as const;

// Each header's column holds its field in base 1.
export const spreadsheetConfigIndexHeaders = {
  startTableColIndexBase0: "Start table column index base 1",
  columnIdRowIdxBase0: "Column ID row index base 1",
  columnGroupHeadingRowIndexBase0: "Column group heading row index base 1",
  actionRowIndexBase0: "Action row index base 1",
  tableHeaderRowIndexBase0: "Table header row index base 1",
} as const;

export function spreadsheetConfigColumnLabel(header: string): string {
  return `Spreadsheet Config column "${header}"`;
}
