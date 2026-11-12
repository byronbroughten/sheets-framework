import { beforeEach, describe, expect, it } from "vitest";

import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  type FakeSheetProperties,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { ColumnConfigOperator } from "./ColumnConfigOperator";
import { TabConfigOperator } from "./TabConfigOperator";
import type { TableConfigOperator } from "./TableConfigOperator";

// The previous configs grant API access, so the fixtures reuse Table IDs the installed configs hold.
const itemTableId = getTableTraitByName("item", "tableId");
const logTableId = getTableTraitByName("log", "tableId");
const tableConfigGid = 210603630;
const widgetGid = 999001;
const gadgetGid = 999003;
const gizmoGid = 999004;

beforeEach(() => {
  stubLogger();
});

function syncedOperator(): TableConfigOperator {
  const columnConfigOperator = ColumnConfigOperator.init();
  columnConfigOperator.ss.fetchAllSheetProperties();
  columnConfigOperator.fetchHeadRows();
  return columnConfigOperator.tableConfigOperator;
}

interface TableFixture {
  name: string;
  columnIds: string[];
  headers?: string[];
}

function stubTables({
  widget,
  gadget,
}: {
  widget: TableFixture;
  gadget?: TableFixture;
}): void {
  stubSheetsService({
    sheets: [
      {
        sheetId: widgetGid,
        title: "Widget",
        rows: buildGridRows({
          0: widget.columnIds,
          3: widget.headers ?? ["ID", "Name"],
        }),
        table: { tableId: itemTableId, name: widget.name, endRowIndex: 5 },
      },
      ...(gadget === undefined
        ? []
        : [
            {
              sheetId: gadgetGid,
              title: "Gadget",
              rows: buildGridRows({
                0: gadget.columnIds,
                3: gadget.headers ?? ["ID", "Name"],
              }),
              table: { tableId: logTableId, name: gadget.name, endRowIndex: 5 },
            },
          ]),
      {
        sheetId: gizmoGid,
        title: "Gizmo",
        rows: buildGridRows({ 0: ["c:gzm:aaa"], 3: ["ID"] }),
        table: { tableId: "gizmo-table", name: "Gizmo", endRowIndex: 5 },
      },
    ],
  });
}

describe("TableConfigOperator.newTableConfigs Table identity and position", () => {
  function stubWidgetTable(): void {
    stubTables({ widget: { name: "Widget orders", columnIds: ["c:itm:aaa"] } });
  }

  it("keys a Table in the previous configs by its live name and records its identity", () => {
    stubWidgetTable();

    expect(syncedOperator().newTableConfigs()).toEqual({
      widgetOrders: {
        tableId: itemTableId,
        tableName: "Widget orders",
        sheetGid: widgetGid,
        idPrefix: "itm",
        hasIdColumn: true,
        hasNameColumn: true,
      },
    });
  });

  it("gives API access only to live Tables the previous configs hold", () => {
    stubWidgetTable();

    expect(syncedOperator().tableIdsApiAccesses()).toEqual([itemTableId]);
  });

  it("writes the entries through makeTableConfigs", () => {
    stubWidgetTable();

    const source = syncedOperator().toFileSource("../makeConfigs");

    expect(source).toContain(
      'import { makeTableConfigs } from "../makeConfigs";',
    );
    expect(source).toContain("export const tableConfigs = makeTableConfigs({");
    expect(source).toContain(
      `"widgetOrders": { "tableId": "${itemTableId}", "tableName": "Widget orders", "sheetGid": 999001, "idPrefix": "itm"`,
    );
  });

  it.each([
    { headers: ["Name"], hasIdColumn: false, hasNameColumn: true },
    { headers: ["ID", "Name"], hasIdColumn: true, hasNameColumn: true },
    { headers: ["ID", "Title"], hasIdColumn: true, hasNameColumn: false },
  ])(
    "emits has-ID $hasIdColumn and has-name $hasNameColumn for header row $headers",
    ({ headers, hasIdColumn, hasNameColumn }) => {
      stubTables({ widget: { name: "Widget", columnIds: [], headers } });

      expect(syncedOperator().newTableConfigs().widget).toMatchObject({
        hasIdColumn,
        hasNameColumn,
      });
    },
  );
});

describe("TableConfigOperator Table keys and prefixes", () => {
  it("fails when two Tables' names give the same key, naming both", () => {
    stubTables({
      widget: { name: "Widget orders", columnIds: ["c:itm:aaa"] },
      gadget: { name: "Widget Orders", columnIds: ["c:log:aaa"] },
    });

    expect(() => syncedOperator().newTableConfigs()).toThrow(
      'Tables "Widget orders" and "Widget Orders" both give the key "widgetOrders".',
    );
  });

  it("keeps a Table's prefix read from its column ID row over the previous one", () => {
    stubTables({ widget: { name: "Widget orders", columnIds: ["c:zzz:aaa"] } });

    expect(syncedOperator().newTableConfigs().widgetOrders?.idPrefix).toBe(
      "zzz",
    );
  });

  it("keeps the previous prefix when the Table holds no IDs", () => {
    stubTables({ widget: { name: "Gizmo", columnIds: [] } });

    expect(syncedOperator().newTableConfigs().gizmo?.idPrefix).toBe("itm");
  });

  it("throws when two Tables share a sampled ID prefix, named by Table name", () => {
    stubTables({
      widget: { name: "Widget", columnIds: ["c:wdg:aaa"] },
      gadget: { name: "Gadget", columnIds: ["c:wdg:bbb"] },
    });

    expect(() => syncedOperator().toFileSource("../makeConfigs")).toThrow(
      /Widget.*Gadget.*"wdg"/,
    );
  });

  it("stops on a sampled prefix that differs from the previous configs", () => {
    stubTables({ widget: { name: "Widget", columnIds: ["c:zzz:aaa"] } });

    expect(() => syncedOperator().validateIdPrefixesUnchanged()).toThrow(
      `Table "${itemTableId}" had ID prefix "itm" and is now "zzz".`,
    );
  });

  it("retires the prefix of a Table gone since the previous configs", () => {
    stubTables({ widget: { name: "Item", columnIds: [] } });

    const operator = syncedOperator();
    expect(operator.retiredIdPrefixes()).toContain("log");
    expect(operator.retiredIdPrefixes()).not.toContain("itm");
  });

  it("writes the retired prefixes through makeRetiredIdPrefixes", () => {
    stubTables({ widget: { name: "Item", columnIds: [] } });

    const source =
      syncedOperator().retiredIdPrefixesFileSource("../makeConfigs");
    expect(source).toContain(
      'import { makeRetiredIdPrefixes } from "../makeConfigs";',
    );
    expect(source).toMatch(
      /export const retiredIdPrefixes = makeRetiredIdPrefixes\(\[.*"log".*\]\);/,
    );
  });
});

describe("TabConfigOperator.newTabConfigs", () => {
  function syncedTabOperator(): TabConfigOperator {
    return new TabConfigOperator(syncedOperator().operatorProps);
  }

  function stubTabs(widgetTab: Partial<FakeSheetProperties>): void {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
        },
        { sheetId: widgetGid, title: "Widget orders", ...widgetTab },
      ],
    });
  }

  it("records each tab's title, colour, frozen counts and hidden flag, keyed by tab name", () => {
    stubTabs({
      tabColorStyle: { rgbColor: { red: 1, blue: 0.5 } },
      frozenRowCount: 4,
      frozenColumnCount: 1,
      hidden: true,
    });

    const tabConfigs = syncedTabOperator().newTabConfigs();
    expect(tabConfigs.widgetOrders).toEqual({
      sheetGid: widgetGid,
      title: "Widget orders",
      color: "#ff0080",
      frozenRows: 4,
      frozenColumns: 1,
      hidden: true,
    });
    expect(tabConfigs.tableConfig).toMatchObject({
      color: null,
      frozenRows: 0,
      frozenColumns: 0,
      hidden: false,
    });
  });

  it("writes the tabs through makeTabConfigs", () => {
    stubTabs({});

    const source = syncedTabOperator().toFileSource("../makeConfigs");
    expect(source).toContain(
      'import { makeTabConfigs } from "../makeConfigs";',
    );
    expect(source).toContain("export const tabConfigs = makeTabConfigs({");
    expect(source).toContain(
      `"widgetOrders": { "sheetGid": ${widgetGid}, "title": "Widget orders"`,
    );
  });

  it("records a theme colour by its name", () => {
    stubTabs({ tabColorStyle: { themeColor: "ACCENT1" } });

    expect(syncedTabOperator().newTabConfigs().widgetOrders?.color).toBe(
      "ACCENT1",
    );
  });

  it("fails when two tab titles give the same tab name, naming both", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
        },
        { sheetId: widgetGid, title: "Widget orders" },
        { sheetId: gadgetGid, title: "Widget Orders" },
      ],
    });

    expect(() => syncedTabOperator().newTabConfigs()).toThrow(
      'Tabs "Widget orders" and "Widget Orders" both give the tab name "widgetOrders".',
    );
  });
});

describe("TableConfigOperator.parseColumnReference", () => {
  beforeEach(() => {
    stubSheetsService({
      sheets: [
        {
          sheetId: widgetGid,
          title: "Rents",
          rows: buildGridRows({
            0: ["c:rnt:aaa", "c:rnt:bbb"],
            3: ["ID", "Tenant"],
          }),
          table: { tableId: itemTableId, name: "Rents", endRowIndex: 5 },
        },
      ],
    });
  });

  it("parses Rents[Tenant] into the live Table's tableId and the header's column ID", () => {
    expect(syncedOperator().parseColumnReference("Rents[Tenant]")).toEqual({
      tableId: itemTableId,
      columnId: "c:rnt:bbb",
    });
  });

  it("fails on a Table the previous configs don't grant API access", () => {
    expect(() =>
      syncedOperator().parseColumnReference("Leases[Tenant]"),
    ).toThrow('names no managed Table "Leases"');
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
