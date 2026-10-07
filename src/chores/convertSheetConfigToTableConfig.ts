import { configSheetFloorSeed } from "../01_SpreadsheetSchema/configSheetFloorSeed";
import { dimensionIds } from "../01_SpreadsheetSchema/dimensionIds";
import type { SpreadsheetRaw } from "../02_SpreadsheetRaw/SpreadsheetRaw";
import type { TableRaw } from "../02_SpreadsheetRaw/TableRaw";
import { retiredSheetConfigTitle } from "../05_Operators/ConfigSheetFloor";
import { Val } from "../utils/Val";
import type { Chore } from "./Chore";

const tableConfigSeed = configSheetFloorSeed.tableConfig;
const headers = {
  sheetGid: "Sheet GID",
  letApiAccess: "Let api access",
  tableId: "Table ID",
} as const;

interface RowConversion {
  rowIndex: number;
  sheetGid: number;
  tableId: string | undefined;
}

// Finds the tab by title, not by generated config, since the configs may still key sheetConfig.
export const convertSheetConfigToTableConfig: Chore = {
  description:
    "Turns the Sheet Config tab into Table Config in place, keeping its GID and every tick: retitles the tab, renames its Table, rewrites Sheet GID as Table ID, and leaves the Table name column to the next config sync's floor. Run once per spreadsheet before gen:configs.",
  action: (ss) => {
    const raw = ss.raw;
    raw.fetchAllSheetProperties();
    const gidByTitle = new Map(
      raw.activeSheetGids.map((gid) => [raw.sheet(gid).title, gid]),
    );
    if (gidByTitle.has(tableConfigSeed.title)) {
      return `Nothing to convert: found a "${tableConfigSeed.title}" tab.`;
    }
    const sheetGid = gidByTitle.get(retiredSheetConfigTitle);
    if (sheetGid === undefined) {
      return `Nothing to convert: no "${retiredSheetConfigTitle}" tab.`;
    }
    const table = raw.tableOnSheet(sheetGid);
    const oldTableName = table.name;
    fetchSheetConfig(raw, table);
    const conversions = rowConversions(raw, table);
    validateTicksHaveOneTable(table, conversions);
    updateSheetGidToTableId(table, conversions);
    table.sheet.updateTitle(tableConfigSeed.title);
    table.updateTableName(tableConfigSeed.liveTableName);
    ss.batchUpdateGSheets();
    return [
      `"${retiredSheetConfigTitle}" → ${tableConfigSeed.title}, keeping GID ${sheetGid}`,
      `Table "${oldTableName}" → ${tableConfigSeed.liveTableName}`,
      ...conversionReport(conversions),
    ].join("; ");
  },
};

function fetchSheetConfig(raw: SpreadsheetRaw, table: TableRaw): void {
  table.headRow("header").gatherFetchFull();
  table.headRow("columnId").gatherFetchFull();
  raw.fetchAllGathered();
  table.columnByHeader(headers.sheetGid).gatherFetchFull();
  table.columnByHeader(headers.letApiAccess).gatherFetchFull();
  raw.fetchAllGathered(true);
}

function rowConversions(raw: SpreadsheetRaw, table: TableRaw): RowConversion[] {
  const tableIdsByGid = tableIdsBySheetGid(raw);
  const gidCol = table.columnByHeader(headers.sheetGid);
  return gidCol.workingCellIndexes.map((rowIndex) => {
    const sheetGid = Number(gidCol.valueOrEmpty(rowIndex));
    const tableIds = tableIdsByGid.get(sheetGid) ?? [];
    return {
      rowIndex,
      sheetGid,
      tableId: tableIds.length === 1 ? tableIds[0] : undefined,
    };
  });
}

function tableIdsBySheetGid(raw: SpreadsheetRaw): Map<number, string[]> {
  return raw.activeTableIds.reduce((byGid, tableId) => {
    const sheetGid = raw.table(tableId).sheetGid;
    const tableIds = byGid.get(sheetGid) ?? [];
    tableIds.push(tableId);
    byGid.set(sheetGid, tableIds);
    return byGid;
  }, new Map<number, string[]>());
}

// A tick on a tab without exactly one Table has no row to carry it, so refuse rather than drop it.
function validateTicksHaveOneTable(
  table: TableRaw,
  conversions: RowConversion[],
): void {
  const tickCol = table.columnByHeader(headers.letApiAccess);
  const lostTick = conversions.find(
    ({ rowIndex, tableId }) =>
      tableId === undefined && tickCol.valueOrEmpty(rowIndex) === true,
  );
  if (lostTick !== undefined) {
    throw new Error(
      `${retiredSheetConfigTitle} ticks GID ${lostTick.sheetGid}, whose tab doesn't hold exactly one Table. Untick it or fix the tab, then rerun.`,
    );
  }
}

function updateSheetGidToTableId(
  table: TableRaw,
  conversions: RowConversion[],
): void {
  const idPrefix = Val.assert(table.profile.idPrefix(), "ID prefix");
  const gidCol = table.columnByHeader(headers.sheetGid);
  const colIndex = gidCol.colIndex;
  table.headRow("header").updateValue(colIndex, headers.tableId);
  table.headRow("columnId").updateValue(colIndex, dimensionIds.col(idPrefix));
  gidCol.updateColumnType("TEXT");
  conversions.forEach(({ rowIndex, tableId }) => {
    gidCol.updateValue(rowIndex, tableId ?? "");
  });
}

function conversionReport(conversions: RowConversion[]): string[] {
  const emptiedGids = conversions
    .filter(({ tableId }) => tableId === undefined)
    .map(({ sheetGid }) => sheetGid);
  const lines = [
    `${headers.sheetGid} → ${headers.tableId} on ${conversions.length} row(s)`,
  ];
  if (emptiedGids.length > 0) {
    lines.push(
      `left ${headers.tableId} empty for unticked GID(s) ${emptiedGids.join(", ")}, whose tab doesn't hold exactly one Table; the sync drops those rows and appends their Tables unticked`,
    );
  }
  return lines;
}
