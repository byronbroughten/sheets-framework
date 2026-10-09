import { Val } from "@byronbroughten/utils/val";

import { configSheetFloorSeed } from "../01_SpreadsheetSchema/configReaders/configSheetFloorSeed";
import { dimensionIds } from "../01_SpreadsheetSchema/dimensionIds";
import { TableOrigin } from "../01_SpreadsheetSchema/TableOrigin";
import type { SpreadsheetRaw } from "../02_SpreadsheetRaw/SpreadsheetRaw";
import type { TableRaw } from "../02_SpreadsheetRaw/TableRaw";
import { retiredSheetConfigTitle } from "../05_Operators/ConfigSheetFloor";
import { retiredSheetGidHeader } from "../05_Operators/ConfigSheetFloor/FloorTabColumnCreator";
import type { Chore } from "./Chore";

const tableConfigSeed = configSheetFloorSeed.tableConfig;
const columnConfigSeed = configSheetFloorSeed.columnConfig;
const headers = {
  sheetGid: retiredSheetGidHeader,
  sheetTitle: "Sheet title",
  letApiAccess: "Let api access",
  tableId: "Table ID",
  tableName: "Table name",
} as const;

interface RowConversion {
  rowIndex: number;
  sheetGid: number;
  tableId: string | undefined;
}

interface TabsToConvert {
  sheetConfig: TableRaw | undefined;
  columnConfig: TableRaw | undefined;
  skipReasons: string[];
}

// Finds the tab by title, not by generated config, since the configs may still key sheetConfig.
export const convertSheetConfigToTableConfig: Chore = {
  description:
    "Turns the Sheet Config tab into Table Config in place, keeping its GID and every tick: retitles the tab, renames its Table, rewrites Sheet GID as Table ID, and leaves the Table name column to the next config sync's floor. On Column Config, rewrites Sheet GID as Table ID and Sheet title as Table name. Skips a tab already converted. Run once per spreadsheet before gen:configs.",
  action: (ss) => {
    const raw = ss.raw;
    raw.fetchAllSheetProperties();
    const tabs = tabsToConvert(raw);
    if (tabs.sheetConfig === undefined && tabs.columnConfig === undefined) {
      return `Nothing to convert: ${tabs.skipReasons.join(", and ")}.`;
    }
    fetchConvertedColumns(raw, tabs);
    const tableIdByGid = tableIdBySheetGid(raw);
    const lines = [
      ...convertSheetConfig(tabs.sheetConfig, tableIdByGid),
      ...convertColumnConfig(raw, tabs.columnConfig, tableIdByGid),
      ...tabs.skipReasons.map((reason) => `nothing else to convert: ${reason}`),
    ];
    ss.batchUpdateGSheets();
    return lines.join("; ");
  },
};

function tabsToConvert(raw: SpreadsheetRaw): TabsToConvert {
  const gidByTitle = new Map(
    raw.activeSheetGids.map((gid) => [raw.sheet(gid).title, gid]),
  );
  const skipReasons: string[] = [];
  let sheetConfig: TableRaw | undefined;
  if (gidByTitle.has(tableConfigSeed.title)) {
    skipReasons.push(`found a "${tableConfigSeed.title}" tab`);
  } else if (gidByTitle.has(retiredSheetConfigTitle)) {
    sheetConfig = tableTitled(raw, gidByTitle, retiredSheetConfigTitle);
  } else {
    skipReasons.push(`no "${retiredSheetConfigTitle}" tab`);
  }
  let columnConfig: TableRaw | undefined;
  if (gidByTitle.has(columnConfigSeed.title)) {
    columnConfig = tableTitled(raw, gidByTitle, columnConfigSeed.title);
  } else {
    skipReasons.push(`no "${columnConfigSeed.title}" tab`);
  }
  fetchHeadRows(raw, [sheetConfig, columnConfig]);
  if (columnConfig?.headRow("header").hasValue(headers.tableId)) {
    skipReasons.push(
      `${columnConfigSeed.title} has a "${headers.tableId}" column`,
    );
    columnConfig = undefined;
  } else if (
    columnConfig !== undefined &&
    !columnConfig.headRow("header").hasValue(headers.sheetGid)
  ) {
    skipReasons.push(
      `${columnConfigSeed.title} has no "${headers.sheetGid}" column`,
    );
    columnConfig = undefined;
  }
  return { sheetConfig, columnConfig, skipReasons };
}

function tableTitled(
  raw: SpreadsheetRaw,
  gidByTitle: Map<string, number>,
  title: string,
): TableRaw {
  return raw.tableOnSheet(Val.assert(gidByTitle.get(title), `"${title}" tab`));
}

function fetchHeadRows(
  raw: SpreadsheetRaw,
  tables: (TableRaw | undefined)[],
): void {
  tables.forEach((table) => {
    table?.headRow("header").gatherFetchFull();
    table?.headRow("columnId").gatherFetchFull();
  });
  raw.fetchAllGathered();
}

function fetchConvertedColumns(
  raw: SpreadsheetRaw,
  { sheetConfig, columnConfig }: TabsToConvert,
): void {
  sheetConfig?.columnByHeader(headers.sheetGid).gatherFetchFull();
  sheetConfig?.columnByHeader(headers.letApiAccess).gatherFetchFull();
  columnConfig?.columnByHeader(headers.sheetGid).gatherFetchFull();
  raw.fetchAllGathered(true);
}

function tableIdBySheetGid(raw: SpreadsheetRaw): Map<number, string> {
  const tableIdsByGid = raw.activeTableIds.reduce((byGid, tableId) => {
    const sheetGid = raw.table(tableId).sheetGid;
    const tableIds = byGid.get(sheetGid) ?? [];
    tableIds.push(tableId);
    byGid.set(sheetGid, tableIds);
    return byGid;
  }, new Map<number, string[]>());
  return [...tableIdsByGid].reduce((byGid, [sheetGid, tableIds]) => {
    const tableId = tableIdOnTab(raw, tableIds);
    if (tableId !== undefined) byGid.set(sheetGid, tableId);
    return byGid;
  }, new Map<number, string>());
}

// A tab with several Tables keeps the one the sheet-centric framework read, at the expected origin.
function tableIdOnTab(
  raw: SpreadsheetRaw,
  tableIds: string[],
): string | undefined {
  if (tableIds.length === 1) return tableIds[0];
  const expected = TableOrigin.expected();
  return tableIds.find((tableId) => raw.table(tableId).origin.equals(expected));
}

function convertSheetConfig(
  table: TableRaw | undefined,
  tableIdByGid: Map<number, string>,
): string[] {
  if (table === undefined) return [];
  const oldTitle = table.sheet.title;
  const oldTableName = table.name;
  const conversions = rowConversions(table, tableIdByGid);
  validateTicksHaveATable(table, conversions);
  updateSheetGidToTableId(table, conversions);
  table.sheet.updateTitle(tableConfigSeed.title);
  table.updateTableName(tableConfigSeed.liveTableName);
  return [
    `"${oldTitle}" → ${tableConfigSeed.title}, keeping GID ${table.sheetGid}`,
    `Table "${oldTableName}" → ${tableConfigSeed.liveTableName}`,
    ...sheetConfigReport(conversions),
  ];
}

function rowConversions(
  table: TableRaw,
  tableIdByGid: Map<number, string>,
): RowConversion[] {
  const gidCol = table.columnByHeader(headers.sheetGid);
  return gidCol.workingCellIndexes.map((rowIndex) => {
    const sheetGid = Number(gidCol.valueOrEmpty(rowIndex));
    return { rowIndex, sheetGid, tableId: tableIdByGid.get(sheetGid) };
  });
}

// A tick on a tab without a Table to carry it has no row, so refuse rather than drop it.
function validateTicksHaveATable(
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
      `${retiredSheetConfigTitle} ticks GID ${lostTick.sheetGid}, ${noTableOnTab()}. Untick it or fix the tab, then rerun.`,
    );
  }
}

function noTableOnTab(): string {
  return "whose tab holds no Table, or several and none at the expected origin";
}

function updateSheetGidToTableId(
  table: TableRaw,
  conversions: RowConversion[],
): void {
  const gidCol = table.columnByHeader(headers.sheetGid);
  const colIndex = gidCol.colIndex;
  table.headRow("header").updateValue(colIndex, headers.tableId);
  table.headRow("columnId").updateValue(colIndex, freshColumnId(table));
  gidCol.updateColumnType("TEXT");
  conversions.forEach(({ rowIndex, tableId }) => {
    gidCol.updateValue(rowIndex, tableId ?? "");
  });
}

function freshColumnId(table: TableRaw): string {
  return dimensionIds.col(Val.assert(table.profile.idPrefix(), "ID prefix"));
}

function sheetConfigReport(conversions: RowConversion[]): string[] {
  const emptiedGids = emptiedSheetGids(conversions);
  const lines = [
    `${headers.sheetGid} → ${headers.tableId} on ${conversions.length} row(s)`,
  ];
  if (emptiedGids.length > 0) {
    lines.push(
      `left ${headers.tableId} empty for unticked GID(s) ${emptiedGids.join(", ")}, ${noTableOnTab()}; the sync drops those rows and appends their Tables unticked`,
    );
  }
  return lines;
}

function emptiedSheetGids(conversions: RowConversion[]): number[] {
  return [
    ...new Set(
      conversions
        .filter(({ tableId }) => tableId === undefined)
        .map(({ sheetGid }) => sheetGid),
    ),
  ];
}

function convertColumnConfig(
  raw: SpreadsheetRaw,
  table: TableRaw | undefined,
  tableIdByGid: Map<number, string>,
): string[] {
  if (table === undefined) return [];
  const conversions = rowConversions(table, tableIdByGid);
  updateSheetGidToTableId(table, conversions);
  updateSheetTitleToTableName(raw, table, conversions);
  return columnConfigReport(conversions);
}

// Written now rather than left to the correction pass, so the tab reads true straight after the chore.
function updateSheetTitleToTableName(
  raw: SpreadsheetRaw,
  table: TableRaw,
  conversions: RowConversion[],
): void {
  const titleCol = table.columnByHeader(headers.sheetTitle);
  const colIndex = titleCol.colIndex;
  table.headRow("header").updateValue(colIndex, headers.tableName);
  table.headRow("columnId").updateValue(colIndex, freshColumnId(table));
  conversions.forEach(({ rowIndex, tableId }) => {
    titleCol.updateValue(
      rowIndex,
      tableId === undefined ? "" : raw.table(tableId).name,
    );
  });
}

function columnConfigReport(conversions: RowConversion[]): string[] {
  const emptiedGids = emptiedSheetGids(conversions);
  const lines = [
    `${columnConfigSeed.title}: ${headers.sheetGid} → ${headers.tableId} and ${headers.sheetTitle} → ${headers.tableName} on ${conversions.length} row(s)`,
  ];
  if (emptiedGids.length > 0) {
    lines.push(
      `left ${columnConfigSeed.title}'s ${headers.tableId} and ${headers.tableName} empty for GID(s) ${emptiedGids.join(", ")}, ${noTableOnTab()}; the sync prunes those rows`,
    );
  }
  return lines;
}
