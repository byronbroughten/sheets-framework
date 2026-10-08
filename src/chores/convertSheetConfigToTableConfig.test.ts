import { beforeEach, describe, expect, it } from "vitest";

import { getTableTraitByName } from "../01_SpreadsheetSchema/tableConfigsTypes";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  type FakeSheetProperties,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import {
  type FakeBodyRow,
  fakeTableSheet,
} from "../testSupport/fakeSheetsService/fakeTableSheet";
import { Val } from "../utils/Val";
import { convertSheetConfigToTableConfig } from "./convertSheetConfigToTableConfig";

const tableConfigGid = getTableTraitByName("tableConfig", "sheetGid");
const columnConfigGid = getTableTraitByName("columnConfig", "sheetGid");
const itemGid = getTableTraitByName("item", "sheetGid");
const twoTableGid = 999001;

const sheetConfigColumns = {
  sheetGid: { columnId: "c:scf:oldGid1", header: "Sheet GID" },
  sheetTitle: { columnId: "c:scf:oldTtl1", header: "Sheet title" },
  letApiAccess: { columnId: "c:scf:apiAcc1", header: "Let api access" },
};

const columnConfigColumns = {
  sheetGid: { columnId: "c:ccf:oldGid1", header: "Sheet GID" },
  tableId: { columnId: "c:ccf:tblId01", header: "Table ID" },
  columnId: { columnId: "c:ccf:colId01", header: "Column ID" },
  sheetTitle: { columnId: "c:ccf:oldTtl1", header: "Sheet title" },
  tableName: { columnId: "c:ccf:tblNam1", header: "Table name" },
  header: { columnId: "c:ccf:header1", header: "Header" },
  emptyValueAllowed: {
    columnId: "c:ccf:empty01",
    header: "Empty value allowed",
  },
};
type ColumnConfigColumn = keyof typeof columnConfigColumns;

const unconvertedColumnConfigNames = [
  "sheetGid",
  "columnId",
  "sheetTitle",
  "header",
  "emptyValueAllowed",
] as const;

beforeEach(() => {
  stubLogger();
});

function withTableName(
  sheet: FakeSheetProperties,
  tableId: string,
  name: string,
): FakeSheetProperties {
  return {
    ...sheet,
    tables: (sheet.tables ?? []).map((table) => ({ ...table, tableId, name })),
  };
}

function sheetConfigTab(
  bodyRows: FakeBodyRow<keyof typeof sheetConfigColumns>[],
): FakeSheetProperties {
  return withTableName(
    fakeTableSheet.build({
      sheetId: tableConfigGid,
      title: "Sheet Config",
      columnConfigs: sheetConfigColumns,
      columnNames: ["sheetGid", "sheetTitle", "letApiAccess"],
      bodyRows,
    }),
    "sheetConfig",
    "sheetConfig",
  );
}

function tableConfigTab(): FakeSheetProperties {
  return {
    sheetId: tableConfigGid,
    title: "Table Config",
    table: { tableId: "sheetConfig", name: "tableConfig", endRowIndex: 5 },
  };
}

function columnConfigTab(
  columnNames: readonly ColumnConfigColumn[],
  bodyRows: FakeBodyRow<ColumnConfigColumn>[],
): FakeSheetProperties {
  return withTableName(
    fakeTableSheet.build({
      sheetId: columnConfigGid,
      title: "Column Config",
      columnConfigs: columnConfigColumns,
      columnNames,
      bodyRows,
    }),
    "columnConfig",
    "columnConfig",
  );
}

function businessTabs(): FakeSheetProperties[] {
  return [
    {
      sheetId: itemGid,
      title: "Item",
      table: { tableId: "item", name: "itemTable", endRowIndex: 5 },
    },
    {
      sheetId: twoTableGid,
      title: "Two Tables",
      tables: [
        { endRowIndex: 5, endColumnIndex: 1 },
        { endRowIndex: 5, startColumnIndex: 2, endColumnIndex: 3 },
      ],
    },
  ];
}

function runChore(sheets: FakeSheetProperties[]) {
  const service = stubSheetsService({ sheets });
  const result = convertSheetConfigToTableConfig.action(
    SpreadsheetNamed.init(),
    { spreadsheetId: "fake" },
  );
  return { ...service, result };
}

describe("convertSheetConfigToTableConfig on Column Config", () => {
  it("rewrites Sheet GID as Table ID and Sheet title as Table name, with fresh column IDs, in one batch update", () => {
    const { grid, batchUpdateCount, result } = runChore([
      tableConfigTab(),
      columnConfigTab(unconvertedColumnConfigNames, [
        {
          sheetGid: itemGid,
          columnId: "c:itm:name001",
          sheetTitle: "Item",
          header: "Name",
          emptyValueAllowed: true,
        },
      ]),
      ...businessTabs(),
    ]);

    const sheet = grid.sheet(columnConfigGid);
    expect(sheet.values({ startRowIndex: 3, endRowIndex: 5 })).toEqual([
      ["Table ID", "Column ID", "Table name", "Header", "Empty value allowed"],
      ["item", "c:itm:name001", "itemTable", "Name", true],
    ]);
    const [tableIdColId, columnIdColId, tableNameColId] = Val.assert(
      sheet.values({ endRowIndex: 1 })[0],
      "column ID row",
    );
    expect(tableIdColId).toMatch(/^c:ccf:/);
    expect(tableIdColId).not.toBe(columnConfigColumns.sheetGid.columnId);
    expect(columnIdColId).toBe(columnConfigColumns.columnId.columnId);
    expect(tableNameColId).toMatch(/^c:ccf:/);
    expect(tableNameColId).not.toBe(columnConfigColumns.sheetTitle.columnId);
    expect(sheet.tables[0]?.columnProperties?.[0]?.columnType).toBe("TEXT");
    expect(batchUpdateCount()).toBe(1);
    expect(result).toContain(
      "Column Config: Sheet GID → Table ID and Sheet title → Table name on 1 row(s)",
    );
  });

  it("leaves Table ID and Table name empty on a row whose tab doesn't hold exactly one Table, and names its GID", () => {
    const { grid, result } = runChore([
      tableConfigTab(),
      columnConfigTab(unconvertedColumnConfigNames, [
        {
          sheetGid: twoTableGid,
          columnId: "c:two:col0001",
          sheetTitle: "Two Tables",
          header: "Left",
          emptyValueAllowed: false,
        },
      ]),
      ...businessTabs(),
    ]);

    expect(grid.sheet(columnConfigGid).bodyValues()).toEqual([
      ["", "c:two:col0001", "", "Left", false],
    ]);
    expect(result).toContain(`GID(s) ${twoTableGid}`);
  });

  it("leaves a Column Config that already has a Table ID column alone, and sends no batch update", () => {
    const { batchUpdateCount, result } = runChore([
      tableConfigTab(),
      columnConfigTab(
        ["tableId", "columnId", "tableName", "header", "emptyValueAllowed"],
        [
          {
            tableId: "item",
            columnId: "c:itm:name001",
            tableName: "itemTable",
            header: "Name",
            emptyValueAllowed: true,
          },
        ],
      ),
      ...businessTabs(),
    ]);

    expect(batchUpdateCount()).toBe(0);
    expect(result).toBe(
      'Nothing to convert: found a "Table Config" tab, and Column Config has a "Table ID" column.',
    );
  });
});

describe("convertSheetConfigToTableConfig on both tabs", () => {
  it("converts an unconverted Sheet Config and Column Config in the same batch update", () => {
    const { grid, batchUpdateCount, result } = runChore([
      sheetConfigTab([
        { sheetGid: itemGid, sheetTitle: "Item", letApiAccess: true },
      ]),
      columnConfigTab(unconvertedColumnConfigNames, [
        {
          sheetGid: itemGid,
          columnId: "c:itm:name001",
          sheetTitle: "Item",
          header: "Name",
          emptyValueAllowed: false,
        },
      ]),
      ...businessTabs(),
    ]);

    expect(grid.sheet(tableConfigGid).title).toBe("Table Config");
    expect(grid.sheet(tableConfigGid).bodyValues()).toEqual([
      ["item", "Item", true],
    ]);
    expect(grid.sheet(columnConfigGid).bodyValues()).toEqual([
      ["item", "c:itm:name001", "itemTable", "Name", false],
    ]);
    expect(batchUpdateCount()).toBe(1);
    expect(result).toContain('"Sheet Config" → Table Config');
    expect(result).toContain("Column Config: Sheet GID → Table ID");
  });
});
