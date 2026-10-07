import type { TableRaw } from "../../02_SpreadsheetRaw/TableRaw";

export function liveColIndex(
  table: TableRaw,
  { columnId, header }: { columnId: string; header: string },
): number | undefined {
  const colIndexes = table.fullTableColIndexes;
  const byId = colIndexes.find(
    (colIndex) =>
      String(table.headRow("columnId").valueOrEmpty(colIndex)) === columnId,
  );
  if (byId !== undefined) return byId;
  return colIndexes.find(
    (colIndex) =>
      String(table.headRow("header").valueOrEmpty(colIndex)) === header,
  );
}
