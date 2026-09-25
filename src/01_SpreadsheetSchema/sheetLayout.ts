// Every sheet's layout and ID conventions; indexes are base 0.
export const sheetLayout = {
  idDelimiter: ":",
  idHeader: "ID",
  nameHeader: "Name",
  startTableColIndex: 0,
  colIdRowIndex: 0,
  colGroupHeadingRowIndex: 1,
  actionRowIndex: 2,
  tableHeaderRowIndex: 3,
} as const;
