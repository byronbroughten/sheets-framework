import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { expectedSheetLayout } from "../testSupport/expectedSheetLayout";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  type FakeSheetProperties,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { fakeTableSheet } from "../testSupport/fakeSheetsService/fakeTableSheet";
import { parseIdNameMapFileSource } from "./parseIdNameMapFileSource";
import { type ConfigRegeneration, SyncCoordinator } from "./SyncCoordinator";
import {
  columnConfigs,
  columnsByHeader,
  type ItemColumnName,
  itemColumns,
  itemGid,
  itemSheet,
  notesGid,
  notesSheet,
  runWithNoAlterations,
  seedFixture,
  sendWithNoAlterations,
  tableConfigsEntry,
} from "./syncCoordinatorTestSupport";

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("SyncCoordinator send with no alterations", () => {
  it("returns every file's source together, reflecting the synced state", () => {
    seedFixture();

    const parsed = sendWithNoAlterations();
    expect(typeof parsed.tableConfigs).toBe("string");
    expect(typeof parsed.columnConfigs).toBe("string");
    expect(parsed.tableConfigs).toContain('"item"');
    expect(parsed.tableConfigs).toContain(
      "export const tableConfigs = makeTableConfigs({",
    );
    expect(parsed.tableConfigs).toContain(
      `"sheetGid": ${itemGid}, "idPrefix": "itm"`,
    );
    expect(typeof parsed.valueConfigs).toBe("string");
    expect(parsed.columnConfigs).toContain(itemColumns.optionalNote.columnId);
  });

  it("writes every make*Configs import from the framework's package name", () => {
    seedFixture();

    const parsed = sendWithNoAlterations();
    function importLine(maker: string): string {
      return `import { ${maker} } from "@byronbroughten/sheets-framework";`;
    }
    expect(parsed.tableConfigs).toContain(importLine("makeTableConfigs"));
    expect(parsed.tabConfigs).toContain(importLine("makeTabConfigs"));
    expect(parsed.conditionalFormatConfigs).toContain(
      importLine("makeConditionalFormatConfigs"),
    );
    expect(parsed.retiredIdPrefixes).toContain(
      importLine("makeRetiredIdPrefixes"),
    );
    expect(parsed.columnConfigs).toContain(importLine("makeColumnConfigs"));
    expect(parsed.valueConfigs).toContain(importLine("makeValueConfigs"));
  });

  it("emits the ID-name map as a type-only file", () => {
    seedFixture();

    const { idNameMap } = sendWithNoAlterations();
    expect(idNameMap).toContain("export interface IdNameMap {");
    expect(idNameMap).not.toMatch(/\b(const|let|function|import)\b/);
  });

  it("maps a Table with API access to its name, tab and columns", () => {
    seedFixture();

    const { tables } = parseIdNameMapFileSource(
      sendWithNoAlterations().idNameMap,
    );
    expect(tables.item).toEqual({
      name: "item",
      sheetGid: itemGid,
      columns: {
        optionalNote: itemColumns.optionalNote.columnId,
        requiredCount: itemColumns.requiredCount.columnId,
        id: itemColumns.id.columnId,
        name: itemColumns.name.columnId,
      },
    });
  });

  it("lists every tab by GID with its name, title and Tables", () => {
    seedFixture({ extraSheets: [notesSheet()] });

    const { tabs } = parseIdNameMapFileSource(
      sendWithNoAlterations().idNameMap,
    );
    expect(tabs[String(itemGid)]).toEqual({
      name: "item",
      title: "Item",
      tableIds: ["item"],
    });
    expect(tabs[String(notesGid)]).toEqual({
      name: "notes",
      title: "Notes",
      tableIds: ["notes"],
    });
  });

  it("emits no Spreadsheet Config file", () => {
    seedFixture();

    const parsed = sendWithNoAlterations();
    expect(Object.keys(parsed)).not.toContain("spreadsheetConfig");
    expect(Object.values(parsed).join("\n")).not.toContain(
      "makeSpreadsheetConfig",
    );
  });

  it("carries the untyped-column summary back, since no run status cell will show it", () => {
    seedFixture();

    const { result } = runWithNoAlterations();

    expect(result.configs.notes).toEqual([
      expect.stringContaining("column(s) across 1 Table(s)"),
    ]);
  });
});

describe("SyncCoordinator send with no alterations API access", () => {
  const logGid = getTableTraitByName("log", "sheetGid");

  function logSheet(): FakeSheetProperties {
    return fakeTableSheet.build({
      sheetId: logGid,
      title: "Log",
      tableId: "log",
      name: "Log",
      columnConfigs: columnConfigs.log,
      columnNames: ["entry", "amount"],
      bodyRows: [{}],
    });
  }

  function tableConfigKeys(files: ConfigRegeneration): string[] {
    return [...files.tableConfigs.matchAll(/^ {2}"([^"]+)": \{/gm)].map(
      (match) => Val.assert(match[1], "Table key"),
    );
  }

  it("gives API access to each live Table the previous configs hold, in their order", () => {
    stubSheetsService({ sheets: [logSheet(), notesSheet(), itemSheet()] });

    expect(tableConfigKeys(sendWithNoAlterations())).toEqual(["item", "log"]);
  });

  it("gives a Table the previous configs don't hold no prefix, no column IDs and no columns in the ID-name map", () => {
    const { grid } = seedFixture({ extraSheets: [notesSheet()] });

    const files = sendWithNoAlterations();

    expect(tableConfigKeys(files)).toEqual(["item"]);
    expect(
      grid.sheet(notesGid).values()[expectedSheetLayout.colIdRowIndex],
    ).toEqual([""]);
    expect(parseIdNameMapFileSource(files.idNameMap).tables.notes).toEqual({
      name: "deskNotes",
      sheetGid: notesGid,
    });
  });
});

describe("SyncCoordinator send with no alterations Empty value allowed", () => {
  function generateItemColumns() {
    const coordinator = SyncCoordinator.init();
    sendWithNoAlterations(coordinator);
    return Val.assert(
      coordinator.columnConfigOperator.newColumnConfigs().item,
      "Item column configs",
    );
  }

  it("carries each column's Empty value allowed forward by column ID, through a header rename", () => {
    const optionalNote: ItemColumnName = "optionalNote";
    seedFixture({
      item: {
        columns: {
          [optionalNote]: {
            columnId: itemColumns.optionalNote.columnId,
            header: "Side note",
          },
        },
      },
    });

    const item = generateItemColumns();

    expect(itemColumns.optionalNote.emptyValueAllowed).toBe(true);
    expect(itemColumns.requiredCount.emptyValueAllowed).toBe(false);
    expect(item.sideNote?.emptyValueAllowed).toBe(true);
    expect(item.requiredCount?.emptyValueAllowed).toBe(false);
  });

  it("gives a new column Empty value allowed false", () => {
    seedFixture({
      item: { columns: { extra: { columnId: "", header: "Extra" } } },
    });

    const item = generateItemColumns();

    expect(item.extra?.columnId).toMatch(/^c:itm:/);
    expect(item.extra?.emptyValueAllowed).toBe(false);
  });
});

describe("SyncCoordinator send with no alterations ID-prefix guard", () => {
  it("stops before writing when a Table's live ID prefix differs from the configs", () => {
    const zzzColumns = Object.fromEntries(
      Object.entries(itemColumns).map(([columnName, column]) => [
        columnName,
        {
          header: column.header,
          columnId: column.columnId.replace("c:itm:", "c:zzz:"),
        },
      ]),
    );
    const { batchUpdateCount } = seedFixture({
      item: { columns: zzzColumns },
    });

    expect(() => sendWithNoAlterations()).toThrow(
      'The regen would change an ID prefix, so it stopped before writing. Table "item" had ID prefix "itm" and is now "zzz".',
    );
    expect(batchUpdateCount()).toBe(0);
  });
});

describe("SyncCoordinator send with no alterations ID prefix", () => {
  const widgetGid = 888001;
  // A Table the previous configs grant, standing in for the fixture's own.
  const widgetTableId = getTableTraitByName("log", "tableId");
  const widgetIdPrefix = getTableTraitByName("log", "idPrefix");

  function widgetSheet(options: {
    title: string;
    tableName?: string;
    columnIds?: readonly string[];
    headers?: readonly string[];
  }): FakeSheetProperties {
    const headers = options.headers ?? ["Name"];
    return fakeTableSheet.build({
      sheetId: widgetGid,
      title: options.title,
      tableId: widgetTableId,
      name: options.tableName ?? options.title,
      columnConfigs: columnsByHeader(headers, options.columnIds),
      columnNames: headers,
      bodyRows: [{}],
    });
  }

  it("samples an existing sheet's prefix from its column IDs, whatever the tab title is now", () => {
    seedFixture({
      extraSheets: [
        widgetSheet({
          title: "Renamed Tab",
          columnIds: [`c:${widgetIdPrefix}:abc1234`],
        }),
      ],
    });

    const parsed = sendWithNoAlterations();
    expect(parsed.tableConfigs).toContain(
      tableConfigsEntry({
        tableKey: "renamedTab",
        tableId: widgetTableId,
        tableName: "Renamed Tab",
        sheetGid: widgetGid,
        idPrefix: widgetIdPrefix,
        hasIdColumn: false,
      }),
    );
  });

  it("mints blank column ID cells on an existing sheet with its prefix", () => {
    seedFixture({
      extraSheets: [
        widgetSheet({
          title: "Widget",
          columnIds: [`c:${widgetIdPrefix}:abc1234`, ""],
          headers: ["Name", "Notes"],
        }),
      ],
    });

    const parsed = sendWithNoAlterations();
    const minted =
      parsed.columnConfigs.match(
        new RegExp(`c:${widgetIdPrefix}:[^"]+`, "g"),
      ) ?? [];
    expect(minted).toHaveLength(2);
    expect(minted).toContain(`c:${widgetIdPrefix}:abc1234`);
  });

  it("stops when one sheet's column IDs carry mixed prefixes, naming the sheet and stray IDs", () => {
    seedFixture({
      extraSheets: [
        widgetSheet({
          title: "Widget",
          columnIds: [`c:${widgetIdPrefix}:abc1234`, "c:gzm:xyz1234"],
          headers: ["Name", "Notes"],
        }),
      ],
    });

    expect(() => sendWithNoAlterations()).toThrow(/Widget.*c:gzm:xyz1234/);
  });

  it("fails when two Tables' names give the same key, naming both", () => {
    seedFixture({
      extraSheets: [widgetSheet({ title: "Widget", tableName: "item" })],
    });

    expect(() => sendWithNoAlterations()).toThrow(
      'Tables "Item" and "item" both give the key "item".',
    );
  });
});
