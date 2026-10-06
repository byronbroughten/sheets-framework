import { beforeEach, describe, expect, it } from "vitest";

import { installedConfigs } from "../01_SpreadsheetSchema/configRegister";
import { TableOrigin } from "../01_SpreadsheetSchema/TableOrigin";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { SheetConfigOperator } from "./SheetConfigOperator";

const { columnConfigs } = installedConfigs();

// The installed columnConfigs' ids, so the fixture matches what production resolves through.
const sc = columnConfigs.sheetConfig;
const sheetConfigGid = 210603630;
const widgetGid = 999001;
const newSheetGid = 999002;
const gadgetGid = 999003;

const sheetConfigColumnIdRow = [
  sc.sheetGid.columnId,
  sc.sheetTitle.columnId,
  sc.letApiAccess.columnId,
];

const existingWidgetConfigRow = [widgetGid, "Widget", true];

beforeEach(() => {
  stubLogger();
});

// newSheetConfigs()/sheetNamesByGid()/toFileSource("../makeConfigs") all read letApiAccess,
// which prepFetchForSync doesn't prep on its own — production code only
// preps it via ColumnConfigOperator.prepFetchWithSheetConfig, so a
// standalone SheetConfigOperator test has to prep it itself.
//
// fetchAllSheetProperties() has to run first, matching
// ConfigCoordinator.syncAndFlushConfigSheets's order — it's what populates
// ss.raw.activeSheetGids (the catalogue walk, and hence
// skipFetchingProperties below) with every live sheet, including ones with
// no Sheet Config row yet.
function syncSheetConfigOperator(operator: SheetConfigOperator): void {
  operator.ss.raw.fetchAllSheetProperties();
  operator.table.prepFetchColumnsFull("letApiAccess");
  operator.prepFetchForSync();
  operator.ss.fetchAllPrepped({ skipFetchingProperties: true });
  operator.syncToSpreadsheet();
}

function syncedOperator(): SheetConfigOperator {
  const operator = SheetConfigOperator.init();
  syncSheetConfigOperator(operator);
  return operator;
}

describe("SheetConfigOperator.newSheetConfigs / toFileSource", () => {
  it("carries forward an existing sheet and appends a brand-new one, excluded until manually enabled", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: sheetConfigGid,
          title: "Sheet Config",
          rows: buildGridRows({
            0: sheetConfigColumnIdRow,
            4: existingWidgetConfigRow,
          }),
          table: { endRowIndex: 5 },
        },
        // Referenced by the existing row above; no "ID" header, so
        // hasIdColumn is emitted false from the header-row sample.
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({ 3: [] }),
          table: { name: "Widget", endRowIndex: 5 },
        },
        // Present in the spreadsheet but with NO existing Sheet Config row.
        {
          sheetId: newSheetGid,
          title: "Brand New Sheet",
          rows: buildGridRows({ 3: [] }),
          table: { name: "Brand New Sheet", endRowIndex: 5 },
        },
      ],
    });

    const operator = SheetConfigOperator.init();
    syncSheetConfigOperator(operator);
    const sheetConfigs = operator.newSheetConfigs();

    expect(sheetConfigs.widget).toEqual({
      sheetGid: widgetGid,
      idPrefix: "wdg",
      hasIdColumn: false,
      hasNameColumn: false,
    });
    // A newly-discovered sheet gets a Sheet Config row appended, but stays
    // excluded from the generated file until a human sets letApiAccess.
    expect(sheetConfigs.brandNewSheet).toBeUndefined();
    expect(operator.table.column("sheetGid").hasValue(newSheetGid)).toBe(true);
  });

  it("resolves sheetGid -> sheetName for a sheet not yet in any deployed config", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: sheetConfigGid,
          title: "Sheet Config",
          rows: buildGridRows({
            0: sheetConfigColumnIdRow,
            // A human already turned on API access for this sheet, but no
            // deploy has run since — this run's own live sync is the only
            // place the mapping exists.
            4: [newSheetGid, "Brand New Sheet", true],
          }),
          table: { endRowIndex: 5 },
        },
        {
          sheetId: newSheetGid,
          title: "Brand New Sheet",
          rows: buildGridRows({ 3: [] }),
          table: { name: "Brand New Sheet", endRowIndex: 5 },
        },
      ],
    });

    const operator = SheetConfigOperator.init();
    syncSheetConfigOperator(operator);

    expect(operator.sheetNamesByGid().get(newSheetGid)).toBe("brandNewSheet");
    expect(operator.newSheetConfigs().brandNewSheet).toEqual({
      sheetGid: newSheetGid,
      idPrefix: "bns",
      hasIdColumn: false,
      hasNameColumn: false,
    });
  });

  // A checkbox nobody has ever touched reads blank, not false.
  it("excludes a sheet whose API-access checkbox has never been ticked", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: sheetConfigGid,
          title: "Sheet Config",
          rows: buildGridRows({
            0: sheetConfigColumnIdRow,
            4: [widgetGid, "Widget", null, "wdg"],
          }),
          table: { endRowIndex: 5 },
        },
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({ 3: [] }),
          table: { name: "Widget", endRowIndex: 5 },
        },
      ],
    });

    const operator = SheetConfigOperator.init();
    syncSheetConfigOperator(operator);

    expect(operator.newSheetConfigs().widget).toBeUndefined();
    expect(operator.sheetGidsApiAccesses()).toEqual([sheetConfigGid]);
  });

  // Every seeded row names a sheet that no longer exists, so the prune reaches the last one.
  it("clears the last stale row rather than deleting it, and syncs past it without throwing", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: sheetConfigGid,
          title: "Sheet Config",
          rows: buildGridRows({
            0: sheetConfigColumnIdRow,
            4: existingWidgetConfigRow,
            5: [newSheetGid, "Gone", true],
          }),
          table: { endRowIndex: 6 },
        },
      ],
    });

    const operator = SheetConfigOperator.init();

    expect(() => syncSheetConfigOperator(operator)).not.toThrow();
    expect(operator.table.row(1).isBlank).toBe(true);
    expect(operator.newSheetConfigs().widget).toBeUndefined();
    expect(operator.newSheetConfigs().sheetConfig).toEqual({
      sheetGid: sheetConfigGid,
      idPrefix: "scf",
      hasIdColumn: false,
      hasNameColumn: false,
    });
  });

  it("assigns an ID prefix from the Table name when a Let api access sheet has no column IDs", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: sheetConfigGid,
          title: "Sheet Config",
          rows: buildGridRows({
            0: sheetConfigColumnIdRow,
            4: [widgetGid, "Widget", true],
          }),
          table: { endRowIndex: 5 },
        },
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({ 3: [] }),
          table: { name: "Widget", endRowIndex: 5 },
        },
      ],
    });

    const operator = SheetConfigOperator.init();
    syncSheetConfigOperator(operator);

    expect(operator.newSheetConfigs().widget).toEqual({
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
          sheetId: sheetConfigGid,
          title: "Sheet Config",
          rows: buildGridRows({
            0: sheetConfigColumnIdRow,
            4: [widgetGid, "Stale Title", true],
          }),
          table: { endRowIndex: 5 },
        },
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({ 3: ["Name"] }),
          table: { name: "Widget", endRowIndex: 5 },
        },
      ],
    });

    const operator = SheetConfigOperator.init();
    syncSheetConfigOperator(operator);

    expect(operator.table.column("sheetTitle").value(0)).toBe("Widget");
    expect(operator.newSheetConfigs().widget?.hasIdColumn).toBe(false);
  });

  it("corrects a draft tab's title from the live tab name without reading its Table", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: sheetConfigGid,
          title: "Sheet Config",
          rows: buildGridRows({
            0: sheetConfigColumnIdRow,
            4: [widgetGid, "Stale Title", false],
          }),
          table: { endRowIndex: 5 },
        },
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({ 3: ["Name"] }),
          table: { endRowIndex: 4 },
        },
      ],
    });

    const operator = SheetConfigOperator.init();
    syncSheetConfigOperator(operator);

    expect(operator.table.column("sheetTitle").value(0)).toBe("Widget");
    expect(operator.newSheetConfigs().widget).toBeUndefined();
  });

  it("emits has-ID true when the described sheet's header row has the ID header", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: sheetConfigGid,
          title: "Sheet Config",
          rows: buildGridRows({
            0: sheetConfigColumnIdRow,
            4: [widgetGid, "Widget", true],
          }),
          table: { endRowIndex: 5 },
        },
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({ 3: ["ID", "Name"] }),
          table: { name: "Widget", endRowIndex: 5 },
        },
      ],
    });

    const operator = SheetConfigOperator.init();
    syncSheetConfigOperator(operator);

    expect(operator.newSheetConfigs().widget?.hasIdColumn).toBe(true);
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
            sheetId: sheetConfigGid,
            title: "Sheet Config",
            rows: buildGridRows({
              0: sheetConfigColumnIdRow,
              4: [widgetGid, "Widget", true],
            }),
            table: { endRowIndex: 5 },
          },
          {
            sheetId: widgetGid,
            title: "Widget",
            rows: buildGridRows({ 3: headers }),
            table: { endRowIndex: 5 },
          },
        ],
      });

      const operator = SheetConfigOperator.init();
      syncSheetConfigOperator(operator);

      expect(operator.newSheetConfigs().widget?.hasNameColumn).toBe(
        hasNameColumn,
      );
    },
  );

  it("throws when two sheets share a sampled ID prefix, named by sheet title", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: sheetConfigGid,
          title: "Sheet Config",
          rows: buildGridRows({
            0: sheetConfigColumnIdRow,
            4: [widgetGid, "Widget", true],
            5: [gadgetGid, "Gadget", true],
          }),
          table: { endRowIndex: 6 },
        },
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({ 0: ["c:wdg:aaa"], 3: [] }),
          table: { endRowIndex: 5 },
        },
        {
          sheetId: gadgetGid,
          title: "Gadget",
          rows: buildGridRows({ 0: ["c:wdg:bbb"], 3: [] }),
          table: { endRowIndex: 5 },
        },
      ],
    });

    const operator = SheetConfigOperator.init();
    syncSheetConfigOperator(operator);

    expect(() => operator.toFileSource("../makeConfigs")).toThrow(
      /Widget.*Gadget.*"wdg"/,
    );
  });
});

describe("SheetConfigOperator.newTableConfigs / toTableConfigsFileSource", () => {
  function stubWidgetTable(letApiAccess: boolean): void {
    stubSheetsService({
      sheets: [
        {
          sheetId: sheetConfigGid,
          title: "Sheet Config",
          rows: buildGridRows({
            0: sheetConfigColumnIdRow,
            4: [widgetGid, "Widget", letApiAccess],
          }),
          table: { name: "Sheet Config", endRowIndex: 5 },
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

  it("keys a managed sheet's Table by its live name and records its identity and position", () => {
    stubWidgetTable(true);

    const operator = SheetConfigOperator.init();
    syncSheetConfigOperator(operator);
    const { headerRowIndex, startColIndex } = TableOrigin.expected();
    const tableConfigs = operator.newTableConfigs();

    expect(Object.keys(tableConfigs).sort()).toEqual([
      "sheetConfig",
      "widgetOrders",
    ]);
    expect(tableConfigs.widgetOrders).toEqual({
      tableId: "widget-table",
      tableName: "Widget orders",
      sheetGid: widgetGid,
      idPrefix: "wdg",
      headerRowIndex,
      startColIndex,
      hasIdColumn: true,
      hasNameColumn: true,
    });
  });

  it("leaves out the Table of a sheet whose API-access checkbox is unticked", () => {
    stubWidgetTable(false);

    const operator = SheetConfigOperator.init();
    syncSheetConfigOperator(operator);

    expect(Object.keys(operator.newTableConfigs())).toEqual(["sheetConfig"]);
  });

  it("writes the entries through makeTableConfigs", () => {
    stubWidgetTable(true);

    const operator = SheetConfigOperator.init();
    syncSheetConfigOperator(operator);
    const source = operator.toTableConfigsFileSource("../makeConfigs");

    expect(source).toContain(
      'import { makeTableConfigs } from "../makeConfigs";',
    );
    expect(source).toContain("export const tableConfigs = makeTableConfigs({");
    expect(source).toContain(
      '"widgetOrders": { "tableId": "widget-table", "tableName": "Widget orders", "sheetGid": 999001, "idPrefix": "wdg"',
    );
  });
});

describe("SheetConfigOperator Table keys and prefixes", () => {
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
          sheetId: sheetConfigGid,
          title: "Sheet Config",
          rows: buildGridRows({
            0: sheetConfigColumnIdRow,
            4: [widgetGid, "Widget", true],
            5: [gadgetGid, "Gadget", true],
          }),
          table: { name: "Sheet Config", endRowIndex: 6 },
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

describe("SheetConfigOperator.parseColumnReference", () => {
  beforeEach(() => {
    stubSheetsService({
      sheets: [
        {
          sheetId: sheetConfigGid,
          title: "Sheet Config",
          rows: buildGridRows({
            0: sheetConfigColumnIdRow,
            4: [widgetGid, "Rents", true],
          }),
          table: { name: "Sheet Config", endRowIndex: 5 },
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
