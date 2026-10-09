import { beforeEach, describe, expect, it } from "vitest";

import { installedConfigs } from "../01_SpreadsheetSchema/configReaders/configRegister";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { tableIdOnTab } from "../testSupport/fakeTableConfigSheet";
import { TableConfigOperator } from "./TableConfigOperator";

const { columnConfigs } = installedConfigs();

// The installed columnConfigs' ids, so the fixture matches what production resolves through.
const tc = columnConfigs.tableConfig;
const tableConfigGid = 210603630;
const widgetGid = 999001;
const newSheetGid = 999002;
const gadgetGid = 999003;
const notesGid = 999004;

const tableConfigTableId = tableIdOnTab(tableConfigGid);
const widgetTableId = tableIdOnTab(widgetGid);
const newSheetTableId = tableIdOnTab(newSheetGid);
const gadgetTableId = tableIdOnTab(gadgetGid);

const tableConfigColumnIdRow = [
  tc.tableId.columnId,
  tc.tableName.columnId,
  tc.sheetTitle.columnId,
  tc.letApiAccess.columnId,
];

const existingWidgetConfigRow = [widgetTableId, "Widget", "Widget", true];

beforeEach(() => {
  stubLogger();
});

// newTableConfigs()/tableKeysByTableId()/toFileSource("../makeConfigs") all read letApiAccess,
// which prepFetchForSync doesn't prep on its own — production code only
// preps it via ColumnConfigOperator.prepFetchWithTableConfig, so a
// standalone TableConfigOperator test has to prep it itself.
//
// fetchAllSheetProperties() has to run first, matching
// ConfigCoordinator.syncAndFlushConfigSheets's order — it's what populates
// ss.raw.activeTableIds (the catalogue walk, and hence
// skipFetchingProperties below) with every live Table, including ones with
// no Table Config row yet.
function syncTableConfigOperator(operator: TableConfigOperator): void {
  operator.ss.raw.fetchAllSheetProperties();
  operator.table.prepFetchColumnsFull("letApiAccess");
  operator.prepFetchForSync();
  operator.ss.fetchAllPrepped({ skipFetchingProperties: true });
  operator.syncToSpreadsheet();
}

function syncedOperator(): TableConfigOperator {
  const operator = TableConfigOperator.init();
  syncTableConfigOperator(operator);
  return operator;
}

function catalogueRows(operator: TableConfigOperator) {
  const col = operator.table.columns(
    "tableId",
    "tableName",
    "sheetTitle",
    "letApiAccess",
  );
  return operator.table.workingRowIndexesWithData.map((rowIndex) => ({
    tableId: col.tableId.value(rowIndex),
    tableName: col.tableName.valueOrEmpty(rowIndex),
    sheetTitle: col.sheetTitle.value(rowIndex),
    letApiAccess: col.letApiAccess.valueOrEmpty(rowIndex),
  }));
}

describe("TableConfigOperator catalogue rows", () => {
  it("lists every Table, new ones unticked, two on one tab as two rows, and gives a tab with no Table no row", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({
            0: tableConfigColumnIdRow,
            4: [null, null, null, null],
          }),
          table: {
            tableId: tableIdOnTab(tableConfigGid),
            name: "tableConfig",
            endRowIndex: 5,
          },
        },
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({ 3: ["ID"] }),
          table: { name: "Widget orders", endRowIndex: 5 },
        },
        {
          sheetId: gadgetGid,
          title: "Gadget",
          rows: buildGridRows({ 3: ["ID"] }),
          tables: [
            { tableId: "gadget-orders", name: "Gadget orders", endRowIndex: 5 },
            {
              tableId: "gadget-stock",
              name: "Gadget stock",
              startRowIndex: 13,
              endRowIndex: 15,
            },
          ],
        },
        { sheetId: notesGid, title: "Notes", rows: buildGridRows({ 3: [] }) },
      ],
    });

    const rows = catalogueRows(syncedOperator());

    expect(rows).toHaveLength(4);
    expect(rows).toEqual(
      expect.arrayContaining([
        {
          tableId: tableConfigTableId,
          tableName: "tableConfig",
          sheetTitle: "Table Config",
          letApiAccess: true,
        },
        {
          tableId: widgetTableId,
          tableName: "Widget orders",
          sheetTitle: "Widget",
          letApiAccess: false,
        },
        {
          tableId: "gadget-orders",
          tableName: "Gadget orders",
          sheetTitle: "Gadget",
          letApiAccess: false,
        },
        {
          tableId: "gadget-stock",
          tableName: "Gadget stock",
          sheetTitle: "Gadget",
          letApiAccess: false,
        },
      ]),
    );
  });

  it("corrects a stale Table name and Sheet title from the live spreadsheet, keeping the tick", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({
            0: tableConfigColumnIdRow,
            4: [widgetTableId, "Old name", "Old title", true],
          }),
          table: {
            tableId: tableIdOnTab(tableConfigGid),
            name: "tableConfig",
            endRowIndex: 5,
          },
        },
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({ 3: ["ID"] }),
          table: { name: "Widget orders", endRowIndex: 5 },
        },
      ],
    });

    const [widgetRow] = catalogueRows(syncedOperator());

    expect(widgetRow).toEqual({
      tableId: widgetTableId,
      tableName: "Widget orders",
      sheetTitle: "Widget",
      letApiAccess: true,
    });
  });
});

describe("TableConfigOperator.newTableConfigs / toFileSource", () => {
  it("carries forward an existing Table and appends a brand-new one, excluded until manually enabled", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({
            0: tableConfigColumnIdRow,
            4: existingWidgetConfigRow,
          }),
          table: {
            tableId: tableIdOnTab(tableConfigGid),
            name: "tableConfig",
            endRowIndex: 5,
          },
        },
        // Referenced by the existing row above; no "ID" header, so
        // hasIdColumn is emitted false from the header-row sample.
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({ 3: [] }),
          table: { name: "Widget", endRowIndex: 5 },
        },
        // Present in the spreadsheet but with NO existing Table Config row.
        {
          sheetId: newSheetGid,
          title: "Brand New Sheet",
          rows: buildGridRows({ 3: [] }),
          table: { name: "Brand New Sheet", endRowIndex: 5 },
        },
      ],
    });

    const operator = TableConfigOperator.init();
    syncTableConfigOperator(operator);
    const tableConfigs = operator.newTableConfigs();

    expect(tableConfigs.widget).toEqual({
      tableId: widgetTableId,
      tableName: "Widget",
      sheetGid: widgetGid,
      idPrefix: "wdg",
      hasIdColumn: false,
      hasNameColumn: false,
    });
    // A newly-discovered Table gets a Table Config row appended, but stays
    // excluded from the generated file until a human sets letApiAccess.
    expect(tableConfigs.brandNewSheet).toBeUndefined();
    expect(operator.table.column("tableId").hasValue(newSheetTableId)).toBe(
      true,
    );
  });

  it("resolves tableId -> tableKey for a Table not yet in any deployed config", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({
            0: tableConfigColumnIdRow,
            // A human already turned on API access for this Table, but no
            // deploy has run since — this run's own live sync is the only
            // place the mapping exists.
            4: [newSheetTableId, "Brand New Sheet", "Brand New Sheet", true],
          }),
          table: {
            tableId: tableIdOnTab(tableConfigGid),
            name: "tableConfig",
            endRowIndex: 5,
          },
        },
        {
          sheetId: newSheetGid,
          title: "Brand New Sheet",
          rows: buildGridRows({ 3: [] }),
          table: { name: "Brand New Sheet", endRowIndex: 5 },
        },
      ],
    });

    const operator = TableConfigOperator.init();
    syncTableConfigOperator(operator);

    expect(operator.tableKeysByTableId().get(newSheetTableId)).toBe(
      "brandNewSheet",
    );
    expect(operator.newTableConfigs().brandNewSheet).toEqual({
      tableId: newSheetTableId,
      tableName: "Brand New Sheet",
      sheetGid: newSheetGid,
      idPrefix: "bns",
      hasIdColumn: false,
      hasNameColumn: false,
    });
  });

  // A checkbox nobody has ever touched reads blank, not false.
  it("excludes a Table whose API-access checkbox has never been ticked", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({
            0: tableConfigColumnIdRow,
            4: [widgetTableId, "Widget", "Widget", null],
          }),
          table: {
            tableId: tableIdOnTab(tableConfigGid),
            name: "tableConfig",
            endRowIndex: 5,
          },
        },
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({ 3: [] }),
          table: { name: "Widget", endRowIndex: 5 },
        },
      ],
    });

    const operator = TableConfigOperator.init();
    syncTableConfigOperator(operator);

    expect(operator.newTableConfigs().widget).toBeUndefined();
    expect(operator.tableIdsApiAccesses()).toEqual([tableConfigTableId]);
  });

  // Every seeded row names a Table that no longer exists, so the prune reaches the last one.
  it("clears the last stale row rather than deleting it, and syncs past it without throwing", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({
            0: tableConfigColumnIdRow,
            4: existingWidgetConfigRow,
            5: [newSheetTableId, "Gone", "Gone", true],
          }),
          table: {
            tableId: tableIdOnTab(tableConfigGid),
            name: "tableConfig",
            endRowIndex: 6,
          },
        },
      ],
    });

    const operator = TableConfigOperator.init();

    expect(() => syncTableConfigOperator(operator)).not.toThrow();
    expect(operator.table.row(1).isBlank).toBe(true);
    expect(operator.newTableConfigs().widget).toBeUndefined();
    expect(operator.newTableConfigs().tableConfig).toEqual({
      tableId: tableConfigTableId,
      tableName: "tableConfig",
      sheetGid: tableConfigGid,
      idPrefix: "scf",
      hasIdColumn: false,
      hasNameColumn: false,
    });
  });

  it("assigns an ID prefix from the Table name when a Let api access Table has no column IDs", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({
            0: tableConfigColumnIdRow,
            4: existingWidgetConfigRow,
          }),
          table: {
            tableId: tableIdOnTab(tableConfigGid),
            name: "tableConfig",
            endRowIndex: 5,
          },
        },
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({ 3: [] }),
          table: { name: "Widget", endRowIndex: 5 },
        },
      ],
    });

    const operator = TableConfigOperator.init();
    syncTableConfigOperator(operator);

    expect(operator.newTableConfigs().widget).toEqual({
      tableId: widgetTableId,
      tableName: "Widget",
      sheetGid: widgetGid,
      idPrefix: "wdg",
      hasIdColumn: false,
      hasNameColumn: false,
    });
  });

  it("emits has-ID from the described sheet's header row and still corrects the sheet title", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({
            0: tableConfigColumnIdRow,
            4: [widgetTableId, "Widget", "Stale Title", true],
          }),
          table: {
            tableId: tableIdOnTab(tableConfigGid),
            name: "tableConfig",
            endRowIndex: 5,
          },
        },
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({ 3: ["Name"] }),
          table: { name: "Widget", endRowIndex: 5 },
        },
      ],
    });

    const operator = TableConfigOperator.init();
    syncTableConfigOperator(operator);

    expect(operator.table.column("sheetTitle").value(0)).toBe("Widget");
    expect(operator.newTableConfigs().widget?.hasIdColumn).toBe(false);
  });

  it("corrects a draft tab's title from the live tab name without reading its Table", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({
            0: tableConfigColumnIdRow,
            4: [widgetTableId, "", "Stale Title", false],
          }),
          table: {
            tableId: tableIdOnTab(tableConfigGid),
            name: "tableConfig",
            endRowIndex: 5,
          },
        },
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({ 3: ["Name"] }),
          table: { endRowIndex: 4 },
        },
      ],
    });

    const operator = TableConfigOperator.init();
    syncTableConfigOperator(operator);

    expect(operator.table.column("sheetTitle").value(0)).toBe("Widget");
    expect(operator.newTableConfigs().widget).toBeUndefined();
  });

  it("emits has-ID true when the described sheet's header row has the ID header", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({
            0: tableConfigColumnIdRow,
            4: existingWidgetConfigRow,
          }),
          table: {
            tableId: tableIdOnTab(tableConfigGid),
            name: "tableConfig",
            endRowIndex: 5,
          },
        },
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({ 3: ["ID", "Name"] }),
          table: { name: "Widget", endRowIndex: 5 },
        },
      ],
    });

    const operator = TableConfigOperator.init();
    syncTableConfigOperator(operator);

    expect(operator.newTableConfigs().widget?.hasIdColumn).toBe(true);
  });

  it.each([
    { headers: ["ID", "Name"], hasNameColumn: true },
    { headers: ["ID", "Title"], hasNameColumn: false },
  ])(
    "emits has-name $hasNameColumn for header row $headers",
    ({ headers, hasNameColumn }) => {
      stubSheetsService({
        sheets: [
          {
            sheetId: tableConfigGid,
            title: "Table Config",
            rows: buildGridRows({
              0: tableConfigColumnIdRow,
              4: existingWidgetConfigRow,
            }),
            table: {
              tableId: tableIdOnTab(tableConfigGid),
              name: "tableConfig",
              endRowIndex: 5,
            },
          },
          {
            sheetId: widgetGid,
            title: "Widget",
            rows: buildGridRows({ 3: headers }),
            table: { name: "Widget", endRowIndex: 5 },
          },
        ],
      });

      const operator = TableConfigOperator.init();
      syncTableConfigOperator(operator);

      expect(operator.newTableConfigs().widget?.hasNameColumn).toBe(
        hasNameColumn,
      );
    },
  );

  it("throws when two Tables share a sampled ID prefix, named by Table name", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({
            0: tableConfigColumnIdRow,
            4: existingWidgetConfigRow,
            5: [gadgetTableId, "", "Gadget", true],
          }),
          table: {
            tableId: tableIdOnTab(tableConfigGid),
            name: "tableConfig",
            endRowIndex: 6,
          },
        },
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({ 0: ["c:wdg:aaa"], 3: [] }),
          table: { name: "Widget", endRowIndex: 5 },
        },
        {
          sheetId: gadgetGid,
          title: "Gadget",
          rows: buildGridRows({ 0: ["c:wdg:bbb"], 3: [] }),
          table: { name: "Gadget", endRowIndex: 5 },
        },
      ],
    });

    const operator = TableConfigOperator.init();
    syncTableConfigOperator(operator);

    expect(() => operator.toFileSource("../makeConfigs")).toThrow(
      /Widget.*Gadget.*"wdg"/,
    );
  });
});

describe("TableConfigOperator.newTableConfigs Table identity and position", () => {
  function stubWidgetTable(letApiAccess: boolean): void {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({
            0: tableConfigColumnIdRow,
            4: ["widget-table", "Widget orders", "Widget", letApiAccess],
          }),
          table: {
            tableId: tableIdOnTab(tableConfigGid),
            name: "tableConfig",
            endRowIndex: 5,
          },
        },
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({ 0: ["c:wdg:aaa"], 3: ["ID", "Name"] }),
          table: {
            tableId: "widget-table",
            name: "Widget orders",
            endRowIndex: 5,
          },
        },
      ],
    });
  }

  it("keys a managed Table by its live name and records its identity", () => {
    stubWidgetTable(true);

    const operator = TableConfigOperator.init();
    syncTableConfigOperator(operator);
    const tableConfigs = operator.newTableConfigs();

    expect(Object.keys(tableConfigs).sort()).toEqual([
      "tableConfig",
      "widgetOrders",
    ]);
    expect(tableConfigs.widgetOrders).toEqual({
      tableId: "widget-table",
      tableName: "Widget orders",
      sheetGid: widgetGid,
      idPrefix: "wdg",
      hasIdColumn: true,
      hasNameColumn: true,
    });
  });

  it("leaves out a Table whose API-access checkbox is unticked", () => {
    stubWidgetTable(false);

    const operator = TableConfigOperator.init();
    syncTableConfigOperator(operator);

    expect(Object.keys(operator.newTableConfigs())).toEqual(["tableConfig"]);
  });

  it("writes the entries through makeTableConfigs", () => {
    stubWidgetTable(true);

    const operator = TableConfigOperator.init();
    syncTableConfigOperator(operator);
    const source = operator.toFileSource("../makeConfigs");

    expect(source).toContain(
      'import { makeTableConfigs } from "../makeConfigs";',
    );
    expect(source).toContain("export const tableConfigs = makeTableConfigs({");
    expect(source).toContain(
      '"widgetOrders": { "tableId": "widget-table", "tableName": "Widget orders", "sheetGid": 999001, "idPrefix": "wdg"',
    );
  });
});

describe("TableConfigOperator Table keys and prefixes", () => {
  interface TableFixture {
    name: string;
    columnIds: string[];
  }
  function stubTwoTables({
    widget,
    gadget,
  }: {
    widget: TableFixture;
    gadget: TableFixture;
  }): void {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({
            0: tableConfigColumnIdRow,
            4: ["widget-table", widget.name, "Widget", true],
            5: ["gadget-table", gadget.name, "Gadget", true],
          }),
          table: {
            tableId: tableIdOnTab(tableConfigGid),
            name: "tableConfig",
            endRowIndex: 6,
          },
        },
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({ 0: widget.columnIds, 3: ["ID", "Name"] }),
          table: { tableId: "widget-table", name: widget.name, endRowIndex: 5 },
        },
        {
          sheetId: gadgetGid,
          title: "Gadget",
          rows: buildGridRows({ 0: gadget.columnIds, 3: ["ID", "Name"] }),
          table: { tableId: "gadget-table", name: gadget.name, endRowIndex: 5 },
        },
      ],
    });
  }

  it("fails when two Tables' names give the same key, naming both", () => {
    stubTwoTables({
      widget: { name: "Widget orders", columnIds: ["c:wdg:aaa"] },
      gadget: { name: "Widget Orders", columnIds: ["c:gdg:aaa"] },
    });

    expect(() => syncedOperator().newTableConfigs()).toThrow(
      'Tables "Widget orders" and "Widget Orders" both give the key "widgetOrders".',
    );
  });

  it("keeps a Table's prefix read from its column ID row over one its name would give", () => {
    stubTwoTables({
      widget: { name: "Widget orders", columnIds: ["c:zzz:aaa"] },
      gadget: { name: "Gadget", columnIds: ["c:gdg:aaa"] },
    });

    expect(syncedOperator().newTableConfigs().widgetOrders?.idPrefix).toBe(
      "zzz",
    );
  });

  it("generates a prefix from the Table's name, not its tab title, when it holds no IDs", () => {
    stubTwoTables({
      widget: { name: "Gizmo", columnIds: [] },
      gadget: { name: "Gadget", columnIds: ["c:gdg:aaa"] },
    });

    expect(syncedOperator().newTableConfigs().gizmo?.idPrefix).toBe("gzm");
  });

  it("keeps a generated prefix unique against one read from another Table", () => {
    stubTwoTables({
      widget: { name: "Gadgets", columnIds: [] },
      gadget: { name: "Gadget", columnIds: ["c:gdg:aaa"] },
    });

    const tableConfigs = syncedOperator().newTableConfigs();
    expect(tableConfigs.gadget?.idPrefix).toBe("gdg");
    expect(tableConfigs.gadgets?.idPrefix).toBe("gdgt");
  });
});

describe("TableConfigOperator.parseColumnReference", () => {
  beforeEach(() => {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({
            0: tableConfigColumnIdRow,
            4: ["rents-table", "Rents", "Rents", true],
          }),
          table: {
            tableId: tableIdOnTab(tableConfigGid),
            name: "tableConfig",
            endRowIndex: 5,
          },
        },
        {
          sheetId: widgetGid,
          title: "Rents",
          rows: buildGridRows({
            0: ["c:rnt:aaa", "c:rnt:bbb"],
            3: ["ID", "Tenant"],
          }),
          table: { tableId: "rents-table", name: "Rents", endRowIndex: 5 },
        },
      ],
    });
  });

  it("parses Rents[Tenant] into the live Table's tableId and the header's column ID", () => {
    expect(syncedOperator().parseColumnReference("Rents[Tenant]")).toEqual({
      tableId: "rents-table",
      columnId: "c:rnt:bbb",
    });
  });

  it("fails on a Table name no managed Table has", () => {
    expect(() =>
      syncedOperator().parseColumnReference("Leases[Tenant]"),
    ).toThrow('names no managed Table "Leases"');
  });

  it("fails on a header the Table doesn't have", () => {
    expect(() =>
      syncedOperator().parseColumnReference("Rents[Landlord]"),
    ).toThrow(/Landlord/);
  });

  it("fails on a reference not of the form Table[Header]", () => {
    expect(() => syncedOperator().parseColumnReference("Rents")).toThrow(
      "not of the form Table[Header]",
    );
  });
});
