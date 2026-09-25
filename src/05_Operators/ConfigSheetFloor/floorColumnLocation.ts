import type { SheetMetaRaw } from "../../02_SpreadsheetRaw/SheetMetaRaw";

export function liveColIndex(
  meta: SheetMetaRaw,
  { columnId, header }: { columnId: string; header: string },
): number | undefined {
  const colIndexes = meta.fullTableColIndexes;
  const byId = colIndexes.find(
    (colIndex) => String(meta.colIdRow.valueOrEmpty(colIndex)) === columnId,
  );
  if (byId !== undefined) return byId;
  return colIndexes.find(
    (colIndex) => String(meta.tableHeaderRow.valueOrEmpty(colIndex)) === header,
  );
}
