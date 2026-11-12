import { Val } from "@byronbroughten/utils/val";
import { beforeEach, describe, expect, it } from "vitest";

import { installedConfigs } from "../01_SpreadsheetSchema/configReaders/configRegister";
import { configSheetFloorSeed } from "../01_SpreadsheetSchema/configReaders/configSheetFloorSeed";
import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { expectedSheetLayout } from "../testSupport/expectedSheetLayout";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  type FakeCell,
  type FakeCellValue,
  type FakeSheetProperties,
  type FakeSheetsService,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { tableIdOnTab } from "../testSupport/fakeTableConfigSheet";
import { parseIdNameMapFileSource } from "../testSupport/parseIdNameMapFileSource";
import {
  ConfigCoordinator,
  type ConfigRegeneration,
} from "./ConfigCoordinator";

type GridRange = GoogleAppsScript.Sheets.Schema.GridRange;

const { columnConfigs } = installedConfigs();
const testSheetGid = getTableTraitByName("item", "sheetGid");
const tableConfigGid = 210603630;
const columnConfigGid = 2034522667;
const headerOnlyTableEndRowIndex = expectedSheetLayout.tableHeaderRowIndex + 1;
const tableHeaderRowIndex = expectedSheetLayout.tableHeaderRowIndex;
const startTableColIndex = expectedSheetLayout.startTableColIndex;
const spreadsheetConfigGid = getTableTraitByName(
  "spreadsheetConfig",
  "sheetGid",
);

const tc = columnConfigs.tableConfig;
const cc = columnConfigs.columnConfig;
const ssc = columnConfigs.spreadsheetConfig;

const sscColumns = [
  "tableMenuSpace",
  "fillRowIdsTimeLastRan",
  "fillRowIdsRunStatus",
  "syncConfigSheetRowsTimeLastRan",
  "syncConfigSheetRowsRunStatus",
] as const;

function floorSeedType(
  tableName: "spreadsheetConfig" | "tableConfig" | "columnConfig",
  header: string,
): string | undefined {
  const column = configSheetFloorSeed[tableName].columns.find(
    (entry) => entry.header === header,
  );
  if (column !== undefined) return column.columnType;
  if (tableName !== "spreadsheetConfig") return undefined;
  return Object.values(configSheetFloorSeed.spreadsheetConfig.endpoints)
    .flatMap((endpoint) => [endpoint.timeLastRan, endpoint.runStatus])
    .find((entry) => entry.header === header)?.columnType;
}

function columnTypesByHeader(
  tableName: "spreadsheetConfig" | "tableConfig" | "columnConfig",
  headers: readonly string[],
): Record<number, string> {
  const types: Record<number, string> = {};
  headers.forEach((header, colIndex) => {
    const columnType = floorSeedType(tableName, header);
    if (columnType !== undefined) types[colIndex] = columnType;
  });
  return types;
}

beforeEach(() => {
  stubLogger();
});

function spreadsheetConfigSheet(): FakeSheetProperties {
  const headers = sscColumns.map((columnName) => ssc[columnName].header);
  return {
    sheetId: spreadsheetConfigGid,
    title: "Spreadsheet Config",
    rows: buildGridRows({
      0: sscColumns.map((columnName) => ssc[columnName].columnId),
      1: headers.map((header) => spreadsheetConfigGroupHeading(header)),
      3: headers,
      4: sscColumns.map((columnName) =>
        columnName === "tableMenuSpace" ? "Not used" : "",
      ),
    }),
    table: {
      tableId: tableIdOnTab(spreadsheetConfigGid),
      name: configSheetFloorSeed.spreadsheetConfig.liveTableName,
      startRowIndex: tableHeaderRowIndex,
      startColumnIndex: startTableColIndex,
      endRowIndex: 5,
      endColumnIndex: headers.length,
      columnTypes: columnTypesByHeader("spreadsheetConfig", headers),
    },
  };
}

function seedFixture(
  options: {
    testColumnId?: string;
    itemHeaderRowIndex?: number;
    extraSheets?: FakeSheetProperties[];
  } = {},
) {
  const testColumnId = options.testColumnId ?? "c:itm:xyz123";
  const itemHeaderRowIndex = options.itemHeaderRowIndex ?? tableHeaderRowIndex;
  return stubSheetsService({
    sheets: [
      spreadsheetConfigSheet(),
      {
        sheetId: tableConfigGid,
        title: "Table Config",
        rows: buildGridRows({
          0: [
            tc.tableId.columnId,
            tc.tableName.columnId,
            tc.sheetTitle.columnId,
            tc.letApiAccess.columnId,
          ],
          3: [
            tc.tableId.header,
            tc.tableName.header,
            tc.sheetTitle.header,
            tc.letApiAccess.header,
          ],
          4: ["item", "", "Item", true],
        }),
        table: {
          tableId: tableIdOnTab(tableConfigGid),
          name: configSheetFloorSeed.tableConfig.liveTableName,
          startRowIndex: tableHeaderRowIndex,
          startColumnIndex: startTableColIndex,
          endRowIndex: 5,
          endColumnIndex: startTableColIndex + 4,
          columnTypes: columnTypesByHeader("tableConfig", [
            tc.tableId.header,
            tc.tableName.header,
            tc.sheetTitle.header,
            tc.letApiAccess.header,
          ]),
        },
      },
      {
        sheetId: columnConfigGid,
        title: "Column Config",
        rows: buildGridRows({
          0: [
            cc.tableId.columnId,
            cc.columnId.columnId,
            cc.tableName.columnId,
            cc.header.columnId,
            cc.emptyValueAllowed.columnId,
          ],
          3: [
            cc.tableId.header,
            cc.columnId.header,
            cc.tableName.header,
            cc.header.header,
            cc.emptyValueAllowed.header,
          ],
          4: [],
        }),
        table: {
          tableId: tableIdOnTab(columnConfigGid),
          name: configSheetFloorSeed.columnConfig.liveTableName,
          startRowIndex: tableHeaderRowIndex,
          startColumnIndex: startTableColIndex,
          endRowIndex: headerOnlyTableEndRowIndex + 1,
          endColumnIndex: startTableColIndex + 5,
          columnTypes: {
            0: "TEXT",
            1: "TEXT",
            2: "TEXT",
            3: "TEXT",
            4: "BOOLEAN",
          },
        },
      },
      {
        sheetId: testSheetGid,
        title: "Item",
        // Row 4 (the top data row) must be present, even blank, and a table
        // range declared, now that
        // ColumnConfigOperator emit samples both (for
        // the live isFormula/valueName facts, the latter via the table's
        // column data-validation rules) for every api-access sheet. The
        // header row (3) needs real text — newColumnConfigs() now throws
        // rather than skips a column still missing one after a sync.
        rows: buildGridRows({
          [itemHeaderRowIndex - 3]: [testColumnId],
          [itemHeaderRowIndex]: ["Some Header"],
          [itemHeaderRowIndex + 1]: [],
        }),
        table: {
          tableId: "item",
          name: "Item",
          startRowIndex: itemHeaderRowIndex,
          endRowIndex: itemHeaderRowIndex + 2,
        },
      },
      ...(options.extraSheets ?? []),
    ],
  });
}

function rowsFromFirstDataRow(
  rows: readonly (readonly FakeCell[])[],
): Record<number, readonly FakeCell[]> {
  return rows.reduce<Record<number, readonly FakeCell[]>>(
    (byIndex, row, offset) => {
      byIndex[expectedSheetLayout.tableHeaderRowIndex + 1 + offset] = row;
      return byIndex;
    },
    {},
  );
}

function spreadsheetConfigGroupHeading(header: string): string {
  const spreadsheetConfigSeed = configSheetFloorSeed.spreadsheetConfig;
  const endpoint = Object.values(spreadsheetConfigSeed.endpoints).find(
    (seededEndpoint) =>
      seededEndpoint.timeLastRan.header === header ||
      seededEndpoint.runStatus.header === header,
  );
  if (endpoint !== undefined) return endpoint.heading;
  return (
    spreadsheetConfigSeed.columns.find((column) => column.header === header)
      ?.columnGroupHeading ?? ""
  );
}

interface TableConfigsEntryFixture {
  tableKey: string;
  tableId: string;
  tableName: string;
  sheetGid: number;
  idPrefix: string;
  hasIdColumn: boolean;
}

// Every fixture Table sits at the expected origin and carries a "Name" header.
function tableConfigsEntry({
  tableKey,
  tableId,
  tableName,
  sheetGid,
  idPrefix,
  hasIdColumn,
}: TableConfigsEntryFixture): string {
  return `"${tableKey}": { "tableId": "${tableId}", "tableName": "${tableName}", "sheetGid": ${sheetGid}, "idPrefix": "${idPrefix}", "hasIdColumn": ${hasIdColumn}, "hasNameColumn": true }`;
}

describe("ConfigCoordinator.generateConfigFiles", () => {
  it("returns every file's source together, reflecting the synced state", () => {
    seedFixture();

    const parsed = ConfigCoordinator.init().generateConfigFiles();
    expect(typeof parsed.tableConfigs).toBe("string");
    expect(typeof parsed.columnConfigs).toBe("string");
    expect(parsed.tableConfigs).toContain('"item"');
    expect(parsed.tableConfigs).toContain(
      "export const tableConfigs = makeTableConfigs({",
    );
    expect(parsed.tableConfigs).toContain(
      `"sheetGid": ${testSheetGid}, "idPrefix": "itm"`,
    );
    expect(typeof parsed.valueConfigs).toBe("string");
    // Gathered into Column Config by the sync, then headed by _updateProgrammaticValues.
    expect(parsed.columnConfigs).toContain("c:itm:xyz123");
  });

  it("writes every make*Configs import from the framework's package name", () => {
    seedFixture();

    const parsed = ConfigCoordinator.init().generateConfigFiles();
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

    const { idNameMap } = ConfigCoordinator.init().generateConfigFiles();
    expect(idNameMap).toContain("export interface IdNameMap {");
    expect(idNameMap).not.toMatch(/\b(const|let|function|import)\b/);
  });

  it("maps every Table's ID to its name and tab, with columns only for a Table with API access", () => {
    const notesGid = 999001;
    seedFixture({
      extraSheets: [
        {
          sheetId: notesGid,
          title: "Notes",
          rows: buildGridRows({ 3: ["Note"], 4: [] }),
          table: { tableId: "notes", name: "Desk Notes", endRowIndex: 5 },
        },
      ],
    });

    const { tables } = parseIdNameMapFileSource(
      ConfigCoordinator.init().generateConfigFiles().idNameMap,
    );
    expect(tables.item).toEqual({
      name: "item",
      sheetGid: testSheetGid,
      columns: { someHeader: "c:itm:xyz123" },
    });
    expect(tables.notes).toEqual({ name: "deskNotes", sheetGid: notesGid });
  });

  it("lists every tab by GID with its name, title and Tables", () => {
    seedFixture();

    const { tabs } = parseIdNameMapFileSource(
      ConfigCoordinator.init().generateConfigFiles().idNameMap,
    );
    expect(tabs[String(testSheetGid)]).toEqual({
      name: "item",
      title: "Item",
      tableIds: ["item"],
    });
    expect(tabs[String(tableConfigGid)]).toMatchObject({
      name: "tableConfig",
      title: "Table Config",
    });
  });

  it("emits no Spreadsheet Config file", () => {
    seedFixture();

    const parsed = ConfigCoordinator.init().generateConfigFiles();
    expect(Object.keys(parsed)).not.toContain("spreadsheetConfig");
    expect(Object.values(parsed).join("\n")).not.toContain(
      "makeSpreadsheetConfig",
    );
  });

  it("carries the untyped-column summary back, since no run status cell will show it", () => {
    seedFixture();

    expect(
      ConfigCoordinator.init().generateConfigFiles().untypedColumnsSummary,
    ).toContain("1 column(s) across 1 Table(s)");
  });
});

describe("ConfigCoordinator.generateConfigFiles ID-prefix guard", () => {
  it("stops before writing when a Table's live ID prefix differs from the configs", () => {
    const { batchUpdateCount } = seedFixture({ testColumnId: "c:zzz:xyz123" });

    expect(() => ConfigCoordinator.init().generateConfigFiles()).toThrow(
      'The regen would change an ID prefix, so it stopped before writing. Table "item" had ID prefix "itm" and is now "zzz".',
    );
    expect(batchUpdateCount()).toBe(0);
  });
});

describe("ConfigCoordinator.generateConfigFiles ID prefix", () => {
  const widgetGid = 888001;
  // A Table the previous configs grant, standing in for the fixture's own.
  const widgetTableId = getTableTraitByName("log", "tableId");
  const widgetIdPrefix = getTableTraitByName("log", "idPrefix");

  function widgetSheet(options: {
    title: string;
    tableId?: string;
    tableName?: string;
    columnIds?: readonly string[];
    headers?: readonly string[];
  }): FakeSheetProperties {
    const headers = options.headers ?? ["Name"];
    const columnIds = options.columnIds ?? headers.map(() => "");
    return {
      sheetId: widgetGid,
      title: options.title,
      rows: buildGridRows({
        0: columnIds,
        3: headers,
        4: [],
      }),
      table: {
        tableId: options.tableId ?? widgetTableId,
        name: options.tableName ?? options.title,
        endRowIndex: 5,
      },
    };
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

    const parsed = ConfigCoordinator.init().generateConfigFiles();
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

    const parsed = ConfigCoordinator.init().generateConfigFiles();
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

    expect(() => ConfigCoordinator.init().generateConfigFiles()).toThrow(
      /Widget.*c:gzm:xyz1234/,
    );
  });

  it("fails when two Tables' names give the same key, naming both", () => {
    seedFixture({
      extraSheets: [widgetSheet({ title: "Widget", tableName: "item" })],
    });

    expect(() => ConfigCoordinator.init().generateConfigFiles()).toThrow(
      'Tables "Item" and "item" both give the key "item".',
    );
  });

  it("gives a Table the configs don't grant no prefix and no column IDs", () => {
    seedFixture({
      extraSheets: [
        widgetSheet({
          title: "Widget",
          tableId: "widget-ungranted",
          columnIds: [""],
        }),
      ],
    });

    const parsed = ConfigCoordinator.init().generateConfigFiles();
    expect(parsed.tableConfigs).not.toContain("widget");
    expect(parsed.columnConfigs).not.toMatch(/c:wdg:/);
  });
});

describe("ConfigCoordinator.generateConfigFiles head-row overlap", () => {
  const recordsGid = 888000222;

  interface RecordsTable {
    name: string;
    headerRowIndex: number;
    startColumnIndex: number;
    endColumnIndex: number;
    endRowIndex: number;
    isManaged: boolean;
  }

  // A managed Table takes a Table ID the previous configs grant.
  const grantedTableIdsByName: Record<string, string> = {
    Leases: getTableTraitByName("log", "tableId"),
    Tenants: getTableTraitByName("runItem", "tableId"),
    Units: getTableTraitByName("computed", "tableId"),
    Owners: getTableTraitByName("dates", "tableId"),
  };

  function recordsTableId(table: RecordsTable): string {
    if (!table.isManaged) return `records-${table.name.toLowerCase()}`;
    return Val.assert(
      grantedTableIdsByName[table.name],
      "granted Table ID of a managed Records Table",
    );
  }

  function recordsSheet(tables: readonly RecordsTable[]): FakeSheetProperties {
    const rows: Record<number, FakeCellValue[]> = {};
    tables.forEach((table) => {
      const headerRow = (rows[table.headerRowIndex] ??= []);
      const topRow = (rows[table.headerRowIndex + 1] ??= []);
      for (
        let col = table.startColumnIndex;
        col < table.endColumnIndex;
        col++
      ) {
        headerRow[col] = `${table.name} ${col}`;
        topRow[col] ??= "";
      }
    });
    return {
      sheetId: recordsGid,
      title: "Records",
      rows: buildGridRows(
        Object.fromEntries(
          Object.entries(rows).map(([rowIndex, row]) => [
            rowIndex,
            Array.from(row, (cell) => cell ?? ""),
          ]),
        ),
      ),
      tables: tables.map((table) => ({
        tableId: recordsTableId(table),
        name: table.name,
        startRowIndex: table.headerRowIndex,
        startColumnIndex: table.startColumnIndex,
        endRowIndex: table.endRowIndex,
        endColumnIndex: table.endColumnIndex,
      })),
    };
  }

  function seedRecords(tables: readonly RecordsTable[]) {
    return seedFixture({ extraSheets: [recordsSheet(tables)] });
  }

  function expectRefused(
    tables: readonly RecordsTable[],
    ...messages: string[]
  ): void {
    const { grid } = seedRecords(tables);
    const configRows = () => ({
      tableConfig: grid.sheet(tableConfigGid).values({
        startRowIndex: 0,
        endRowIndex: 30,
        endColumnIndex: 10,
      }),
      columnConfig: grid.sheet(columnConfigGid).values({
        startRowIndex: 0,
        endRowIndex: 30,
        endColumnIndex: 10,
      }),
    });
    const before = configRows();

    let parsed: ConfigRegeneration | undefined;
    expect(() => {
      parsed = ConfigCoordinator.init().generateConfigFiles();
    }).toThrow(messages.join(" "));
    expect(parsed).toBeUndefined();
    expect(configRows()).toEqual(before);
  }

  // Rows 4–9 are its body, 0-based.
  const leases: RecordsTable = {
    name: "Leases",
    headerRowIndex: 3,
    startColumnIndex: 1,
    endColumnIndex: 4,
    endRowIndex: 10,
    isManaged: true,
  };

  it("refuses a Table stacked under another and starting left of it, its head rows on the other's last body row", () => {
    expectRefused(
      [
        leases,
        {
          name: "Tenants",
          headerRowIndex: 12,
          startColumnIndex: 0,
          endColumnIndex: 3,
          endRowIndex: 14,
          isManaged: true,
        },
      ],
      `Table "Tenants" on sheet "Records" has head rows in rows 10–12 that sit on Table "Leases". Move "Tenants" so that the 3 rows above its header sit clear of "Leases".`,
    );
  });

  it("refuses a Table whose head rows sit on another managed Table's head rows", () => {
    expectRefused(
      [
        {
          name: "Leases",
          headerRowIndex: 3,
          startColumnIndex: 0,
          endColumnIndex: 3,
          endRowIndex: 5,
          isManaged: true,
        },
        {
          name: "Tenants",
          headerRowIndex: 5,
          startColumnIndex: 0,
          endColumnIndex: 3,
          endRowIndex: 7,
          isManaged: true,
        },
      ],
      `Table "Tenants" on sheet "Records" has head rows in rows 3–5 that sit on Table "Leases". Move "Tenants" so that the 3 rows above its header sit clear of "Leases".`,
    );
  });

  it("refuses a managed Table whose head rows sit on a Table the app doesn't manage", () => {
    expectRefused(
      [
        { ...leases, isManaged: false },
        {
          name: "Tenants",
          headerRowIndex: 12,
          startColumnIndex: 0,
          endColumnIndex: 3,
          endRowIndex: 14,
          isManaged: true,
        },
      ],
      `Table "Tenants" on sheet "Records" has head rows in rows 10–12 that sit on Table "Leases". Move "Tenants" so that the 3 rows above its header sit clear of "Leases".`,
    );
  });

  it("lists two overlaps on one sheet in one message", () => {
    expectRefused(
      [
        {
          name: "Leases",
          headerRowIndex: 3,
          startColumnIndex: 0,
          endColumnIndex: 2,
          endRowIndex: 10,
          isManaged: true,
        },
        {
          name: "Units",
          headerRowIndex: 3,
          startColumnIndex: 3,
          endColumnIndex: 5,
          endRowIndex: 10,
          isManaged: true,
        },
        {
          name: "Tenants",
          headerRowIndex: 12,
          startColumnIndex: 0,
          endColumnIndex: 2,
          endRowIndex: 14,
          isManaged: true,
        },
        {
          name: "Owners",
          headerRowIndex: 11,
          startColumnIndex: 3,
          endColumnIndex: 5,
          endRowIndex: 13,
          isManaged: true,
        },
      ],
      `Table "Tenants" on sheet "Records" has head rows in rows 10–12 that sit on Table "Leases". Move "Tenants" so that the 3 rows above its header sit clear of "Leases".`,
      `Table "Owners" on sheet "Records" has head rows in rows 9–11 that sit on Table "Units". Move "Owners" so that the 3 rows above its header sit clear of "Units".`,
    );
  });

  it("refuses a granted Table stacked clear of another but below the header zone, naming it", () => {
    seedRecords([
      leases,
      {
        name: "Tenants",
        headerRowIndex: 13,
        startColumnIndex: 0,
        endColumnIndex: 3,
        endRowIndex: 15,
        isManaged: true,
      },
    ]);

    expect(() => ConfigCoordinator.init().generateConfigFiles()).toThrowError(
      /^Table "Tenants" on "Records" \(gid \d+\) must have its header row on row 4\.$/,
    );
  });

  it("refuses a recorded Table moved below the header zone with generation's own fix", () => {
    seedFixture({ itemHeaderRowIndex: tableHeaderRowIndex + 2 });

    expect(() => ConfigCoordinator.init().generateConfigFiles()).toThrowError(
      /^Table "Item" on "Item" \(gid \d+\) must have its header row on row 4\.$/,
    );
  });

  it("leaves an ungranted Table below the header zone to the operator", () => {
    seedRecords([
      leases,
      {
        name: "Tenants",
        headerRowIndex: 13,
        startColumnIndex: 0,
        endColumnIndex: 3,
        endRowIndex: 15,
        isManaged: false,
      },
    ]);

    const parsed = ConfigCoordinator.init().generateConfigFiles();
    expect(parsed.tableConfigs).toContain('"leases"');
    expect(parsed.tableConfigs).not.toContain('"tenants"');
  });

  it("allows two Tables side by side, touching", () => {
    seedRecords([
      {
        name: "Leases",
        headerRowIndex: 3,
        startColumnIndex: 0,
        endColumnIndex: 2,
        endRowIndex: 6,
        isManaged: true,
      },
      {
        name: "Tenants",
        headerRowIndex: 3,
        startColumnIndex: 2,
        endColumnIndex: 4,
        endRowIndex: 5,
        isManaged: true,
      },
    ]);

    const parsed = ConfigCoordinator.init().generateConfigFiles();
    expect(parsed.tableConfigs).toContain('"tenants"');
    expect(parsed.tableConfigs).toContain('"leases"');
  });
});

describe("ConfigCoordinator conditional-format rule read", () => {
  const rulesGid = 777000222;
  const widgetGid = 777000333;
  // A Table the previous configs grant, so its rules go under its Table key.
  const widgetTableId = getTableTraitByName("log", "tableId");
  const separateRuleGet = "sheets(properties(sheetId),conditionalFormats)";

  function backgroundRule(
    range: GridRange,
  ): GoogleAppsScript.Sheets.Schema.ConditionalFormatRule {
    return {
      ranges: [range],
      booleanRule: {
        condition: {
          type: "CUSTOM_FORMULA",
          values: [{ userEnteredValue: "=$B5>1" }],
        },
        format: { backgroundColor: { red: 1, green: 0, blue: 0 } },
      },
    };
  }

  function rulesSheet(): FakeSheetProperties {
    return {
      sheetId: rulesGid,
      title: "Rules",
      rows: buildGridRows({}),
      conditionalFormats: [
        backgroundRule({
          sheetId: rulesGid,
          startRowIndex: 4,
          endRowIndex: 8,
          startColumnIndex: 1,
          endColumnIndex: 2,
        }),
      ],
    };
  }

  it("brings every sheet's rules into Raw on the regen's plain get, with no separate rule get", () => {
    const service = seedFixture({ extraSheets: [rulesSheet()] });
    const coordinator = ConfigCoordinator.init();

    coordinator.generateConfigFiles();

    expect(coordinator.ss.raw.sheet(rulesGid).conditionalFormatRules()).toEqual(
      [
        {
          kind: "boolean",
          ranges: [
            {
              sheetId: rulesGid,
              startRowIndex: 4,
              endRowIndex: 8,
              startColumnIndex: 1,
              endColumnIndex: 2,
            },
          ],
          condition: { type: "CUSTOM_FORMULA", formula: "=$B5>1" },
          format: { backgroundColor: { red: 1, green: 0, blue: 0 } },
        },
      ],
    );
    expect(service.getCalls.map(({ fields }) => fields)).not.toContain(
      separateRuleGet,
    );
  });

  it("writes a Table's rules under its Table key and a tab's under its tab name, each by its rule handle", () => {
    seedFixture({
      extraSheets: [
        rulesSheet(),
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({ 3: ["Name", "Score"], 4: [] }),
          table: { tableId: widgetTableId, name: "Widget", endRowIndex: 5 },
          conditionalFormats: [
            backgroundRule({
              sheetId: widgetGid,
              startRowIndex: 4,
              endRowIndex: 5,
              startColumnIndex: 1,
              endColumnIndex: 2,
            }),
          ],
        },
      ],
    });

    const { conditionalFormatConfigs } =
      ConfigCoordinator.init().generateConfigFiles();

    expect(conditionalFormatConfigs).toContain(
      "export const conditionalFormatConfigs = makeConditionalFormatConfigs({",
    );
    expect(conditionalFormatConfigs).toContain(
      `  "tables": {\n    "widget": {\n      "score | =\${score}>1 | background #ff0000": { "kind": "boolean", "tableId": "${widgetTableId}", "columns": ["score"], "condition": {"type":"CUSTOM_FORMULA","formula":"=\${score}>1"}, "format": {"backgroundColor":"#ff0000"} }\n    }\n  },`,
    );
    expect(conditionalFormatConfigs).toContain(
      `  "tabs": {\n    "rules": {\n      "B5:B8 | =$B5>1 | background #ff0000": { "kind": "boolean", "sheetGid": ${rulesGid}, "ranges": ["B5:B8"], "condition": {"type":"CUSTOM_FORMULA","formula":"=$B5>1"}, "format": {"backgroundColor":"#ff0000"} }\n    }\n  }\n});`,
    );
  });
});

describe("ConfigCoordinator.generateConfigFiles value sources", () => {
  const listGid = 889001;
  const ledgerGid = 889002;
  const listPaymentGid = 889004;

  interface FixtureColumn {
    header: string;
    values?: readonly FakeCell[];
    columnType?: string;
    validationValues?: string[];
    validationConditionType?: string;
    actionRowCell?: FakeCellValue;
    groupHeading1?: string;
  }

  function tableSheet(props: {
    sheetId: number;
    title: string;
    columns: readonly FixtureColumn[];
  }): FakeSheetProperties {
    const { columns } = props;
    const rowCount = Math.max(
      1,
      ...columns.map((column) => column.values?.length ?? 0),
    );
    const dataRows: (readonly FakeCell[])[] = Array.from(
      { length: rowCount },
      (_, rowOffset) =>
        columns.map((column) => column.values?.[rowOffset] ?? null),
    );
    return {
      sheetId: props.sheetId,
      title: props.title,
      rows: buildGridRows({
        0: columns.map(() => ""),
        1: columns.map((column) => column.groupHeading1 ?? null),
        2: columns.map((column) => column.actionRowCell ?? null),
        3: columns.map((column) => column.header),
        ...rowsFromFirstDataRow(dataRows),
      }),
      table: {
        name: props.title,
        endRowIndex: headerOnlyTableEndRowIndex + rowCount,
        columnTypes: byColIndex(columns, (column) => column.columnType),
        columnValidationValues: byColIndex(
          columns,
          (column) => column.validationValues,
        ),
        columnValidationConditionTypes: byColIndex(
          columns,
          (column) => column.validationConditionType,
        ),
      },
    };
  }

  function byColIndex<V>(
    columns: readonly FixtureColumn[],
    read: (column: FixtureColumn) => V | undefined,
  ): Record<number, V> {
    return columns.reduce<Record<number, V>>((byIndex, column, colIndex) => {
      const value = read(column);
      if (value !== undefined) byIndex[colIndex] = value;
      return byIndex;
    }, {});
  }

  function seedValueSources(props: {
    apiTables: readonly FakeSheetProperties[];
    otherTables?: readonly FakeSheetProperties[];
  }): FakeSheetsService {
    const { apiTables, otherTables = [] } = props;
    return seedFixture({
      extraSheets: [...apiTables.map(withGrantedTableId), ...otherTables],
    });
  }

  // The previous configs grant API access, so an API-access fixture takes a Table ID they hold.
  function withGrantedTableId(sheet: FakeSheetProperties): FakeSheetProperties {
    const grantedTableIdsByGid = new Map([
      [listGid, getTableTraitByName("log", "tableId")],
      [ledgerGid, getTableTraitByName("runItem", "tableId")],
      [listPaymentGid, getTableTraitByName("computed", "tableId")],
    ]);
    return {
      ...sheet,
      table: {
        ...Val.assert(sheet.table, "API-access fixture Table"),
        tableId: Val.assert(
          grantedTableIdsByGid.get(sheet.sheetId),
          "granted Table ID of an API-access fixture",
        ),
      },
    };
  }

  function ledgerSheet(columns: readonly FixtureColumn[]): FakeSheetProperties {
    return tableSheet({ sheetId: ledgerGid, title: "Ledger", columns });
  }

  function listSheet(columns: readonly FixtureColumn[]): FakeSheetProperties {
    return tableSheet({ sheetId: listGid, title: "List", columns });
  }

  function generate() {
    const coordinator = ConfigCoordinator.init();
    const files = coordinator.generateConfigFiles();
    return {
      valueConfigs: emittedValueConfigs(files),
      ledgerValueName: (columnName: string) =>
        coordinator.columnConfigOperator.newColumnConfigs().ledger?.[columnName]
          ?.valueName,
    };
  }

  // A source fetch reads one whole column; row samples span the Table.
  function fullyFetchedListColumnIndexes(service: FakeSheetsService): number[] {
    const indexes = service.getByDataFilterCalls.flatMap((resource) =>
      (
        (resource as { dataFilters?: { gridRange?: GridRange }[] })
          .dataFilters ?? []
      )
        .map(({ gridRange }) => gridRange)
        .filter(
          (range): range is GridRange =>
            range?.sheetId === listGid &&
            range.startRowIndex === 0 &&
            (range.endColumnIndex ?? 0) - (range.startColumnIndex ?? 0) === 1,
        )
        .map((range) => Val.assert(range.startColumnIndex, "start column")),
    );
    return [...new Set(indexes)].sort();
  }

  function emittedValueConfigs(files: ConfigRegeneration): unknown {
    const match = files.valueConfigs.match(/makeValueConfigs\(([\s\S]*)\);/);
    return JSON.parse(Val.assert(match?.[1], "emitted value configs"));
  }

  it("types a dropdown from a Let api access Table column as that column's fixed list", () => {
    seedValueSources({
      apiTables: [
        listSheet([
          { header: "Payment status", values: ["Paid", "Unpaid", "Late"] },
        ]),
        ledgerSheet([
          {
            header: "Status",
            values: ["Paid"],
            validationValues: ["=List[Payment status]"],
          },
        ]),
      ],
    });

    const { valueConfigs, ledgerValueName } = generate();

    expect(ledgerValueName("status")).toBe("listPaymentStatus");
    expect(valueConfigs).toEqual({
      listPaymentStatus: ["Paid", "Unpaid", "Late"],
    });
  });

  function seedNameAndIdSources(): FakeSheetsService {
    return seedValueSources({
      apiTables: [
        listSheet([
          { header: "ID", values: ["r:lst:a1"] },
          { header: "Name", values: ["Unit 2B"], columnType: "DOUBLE" },
          { header: "Kind", values: ["Flat"] },
        ]),
        ledgerSheet([
          { header: "Unit", validationValues: ["=List[Name]"] },
          { header: "Unit row", validationValues: ["=List[ID]"] },
          { header: "Kind", validationValues: ["=List[Kind]"] },
        ]),
      ],
    });
  }

  it("emits a dropdown from a Table's Name or ID column as a live reference, whatever its type", () => {
    seedNameAndIdSources();

    const { valueConfigs, ledgerValueName } = generate();

    expect(
      ["unit", "unitRow", "kind"].map((columnName) =>
        ledgerValueName(columnName),
      ),
    ).toEqual(["listName", "listId", "listKind"]);
    expect(valueConfigs).toEqual({
      listName: { kind: "liveReference" },
      listId: { kind: "liveReference" },
      listKind: ["Flat"],
    });
  });

  it("fetches no Name or ID source column, only the fixed-list ones", () => {
    const service = seedNameAndIdSources();

    generate();

    expect(fullyFetchedListColumnIndexes(service)).toEqual([2]);
  });

  it("keeps a one-word camel-case header's casing in the value name", () => {
    seedValueSources({
      apiTables: [
        listSheet([{ header: "CapEx", values: ["Roof"] }]),
        ledgerSheet([{ header: "CapEx", validationValues: ["=List[CapEx]"] }]),
      ],
    });

    const { valueConfigs, ledgerValueName } = generate();

    expect(ledgerValueName("capEx")).toBe("listCapEx");
    expect(valueConfigs).toEqual({ listCapEx: ["Roof"] });
  });

  it("types every member by its source's declared type, numbers for number and date", () => {
    seedValueSources({
      apiTables: [
        listSheet([
          { header: "Rent", values: [950, 1200], columnType: "DOUBLE" },
          { header: "Due date", values: [45000, 45031], columnType: "DATE" },
          { header: "Unit", values: ["2B", 101], columnType: "TEXT" },
        ]),
        ledgerSheet([
          { header: "Rent", validationValues: ["=List[Rent]"] },
          { header: "Due", validationValues: ["=List[Due date]"] },
          { header: "Unit", validationValues: ["=List[Unit]"] },
        ]),
      ],
    });

    const { valueConfigs } = generate();

    expect(valueConfigs).toEqual({
      listRent: [950, 1200],
      listDueDate: [45000, 45031],
      listUnit: ["2B", "101"],
    });
  });

  it("fails the config sync when a declared source holds a member of another primitive, naming the source", () => {
    seedValueSources({
      apiTables: [
        listSheet([
          { header: "Rent", values: [950, "TBD"], columnType: "DOUBLE" },
        ]),
        ledgerSheet([{ header: "Rent", validationValues: ["=List[Rent]"] }]),
      ],
    });

    expect(() => ConfigCoordinator.init().generateConfigFiles()).toThrow(
      'Value source "List[Rent]" declares number members, but holds "TBD".',
    );
  });

  it("keeps each member's own primitive on an untyped source", () => {
    seedValueSources({
      apiTables: [
        listSheet([{ header: "Score", values: ["N/A", 3, true] }]),
        ledgerSheet([{ header: "Score", validationValues: ["=List[Score]"] }]),
      ],
    });

    const { valueConfigs } = generate();

    expect(valueConfigs).toEqual({ listScore: ["N/A", 3, true] });
  });

  it("drops blanks and collapses duplicates to their first position, keeping sheet order", () => {
    seedValueSources({
      apiTables: [
        listSheet([
          {
            header: "Payment status",
            values: ["Unpaid", null, "Paid", "Unpaid", "", "Late", "Paid"],
          },
        ]),
        ledgerSheet([
          { header: "Status", validationValues: ["=List[Payment status]"] },
        ]),
      ],
    });

    const { valueConfigs } = generate();

    expect(valueConfigs).toEqual({
      listPaymentStatus: ["Unpaid", "Paid", "Late"],
    });
  });

  it("gives every dropdown on one source the one value name", () => {
    seedValueSources({
      apiTables: [
        listSheet([{ header: "Payment status", values: ["Paid", "Unpaid"] }]),
        ledgerSheet([
          { header: "Status", validationValues: ["=List[Payment status]"] },
          {
            header: "Deposit status",
            validationValues: ["=List[Payment status]"],
          },
        ]),
      ],
    });

    const { valueConfigs, ledgerValueName } = generate();

    expect(ledgerValueName("status")).toBe("listPaymentStatus");
    expect(ledgerValueName("depositStatus")).toBe("listPaymentStatus");
    expect(valueConfigs).toEqual({ listPaymentStatus: ["Paid", "Unpaid"] });
  });

  it("leaves a dropdown untyped, and in the summary, unless it names a column of a Let api access Table", () => {
    seedValueSources({
      apiTables: [
        listSheet([{ header: "Payment status", values: ["Paid", "Unpaid"] }]),
        ledgerSheet([
          {
            header: "Non-api",
            values: ["Paid"],
            validationValues: ["=Hidden[Payment status]"],
          },
          {
            header: "Unknown header",
            values: ["Paid"],
            validationValues: ["=List[Nope]"],
          },
          {
            header: "A1 range",
            values: ["Paid"],
            validationValues: ["=List!$A$5:$A$9"],
          },
          {
            header: "Item list",
            values: ["Paid"],
            validationValues: ["Paid", "Unpaid"],
          },
        ]),
      ],
      otherTables: [
        tableSheet({
          sheetId: 889003,
          title: "Hidden",
          columns: [{ header: "Payment status", values: ["Paid"] }],
        }),
      ],
    });
    const logger = stubLogger();

    const coordinator = ConfigCoordinator.init();
    const files = coordinator.generateConfigFiles();

    expect(emittedValueConfigs(files)).toEqual({});
    const ledger = coordinator.columnConfigOperator.newColumnConfigs().ledger;
    expect(
      ["nonApi", "unknownHeader", "a1Range", "itemList"].map(
        (columnName) => ledger?.[columnName]?.valueName,
      ),
    ).toEqual(["string", "string", "string", "string"]);
    expect(files.untypedColumnsSummary).toContain("column(s) across");
    expect(logger.log).toHaveBeenCalledWith(
      'Untyped columns on Table "Ledger" (4): Non-api, Unknown header, A1 range, Item list',
    );
  });

  it("fails the config sync when a source's value name is a framework value name, naming both", () => {
    seedValueSources({
      apiTables: [
        tableSheet({
          sheetId: listGid,
          title: "Number",
          columns: [{ header: "#", values: [1, 2] }],
        }),
        ledgerSheet([{ header: "Pick", validationValues: ["=Number[#]"] }]),
      ],
    });

    expect(() => ConfigCoordinator.init().generateConfigFiles()).toThrow(
      'Value source "Number[#]" takes the value name "number", which is already the framework value name "number".',
    );
  });

  it("fails the config sync when two sources take one value name, naming both", () => {
    seedValueSources({
      apiTables: [
        listSheet([{ header: "Payment status", values: ["Paid"] }]),
        tableSheet({
          sheetId: listPaymentGid,
          title: "List Payment",
          columns: [{ header: "Status", values: ["Open"] }],
        }),
        ledgerSheet([
          { header: "Status", validationValues: ["=List[Payment status]"] },
          { header: "State", validationValues: ["=List Payment[Status]"] },
        ]),
      ],
    });

    expect(() => ConfigCoordinator.init().generateConfigFiles()).toThrow(
      'Value sources "List[Payment status]" and "List Payment[Status]" both take the value name "listPaymentStatus".',
    );
  });

  describe("the baseline each column records", () => {
    function generateLedgerColumns() {
      const coordinator = ConfigCoordinator.init();
      const files = coordinator.generateConfigFiles();
      return {
        files,
        ledger: Val.assert(
          coordinator.columnConfigOperator.newColumnConfigs().ledger,
          "Ledger column configs",
        ),
      };
    }

    it("records formula text in brace form, naming its own Table's columns bare and another Table's in full", () => {
      seedValueSources({
        apiTables: [
          listSheet([{ header: "Payment status", values: ["Paid"] }]),
          ledgerSheet([
            { header: "Amount", values: [10] },
            {
              header: "Double",
              values: [
                {
                  value: 20,
                  isFormula: true,
                  formula: "=SINGLE(Ledger[Amount])*2",
                },
              ],
            },
            {
              header: "Statuses",
              values: [
                {
                  value: 1,
                  isFormula: true,
                  formula: '=COUNTA(List[Payment status])&"Ledger[Amount]"',
                },
              ],
            },
          ]),
        ],
      });

      const { files, ledger } = generateLedgerColumns();

      expect(ledger.amount?.formula).toBeNull();
      expect(ledger.double?.formula).toBe("=${amount}*2");
      expect(ledger.statuses?.formula).toBe(
        '=COUNTA(${list[paymentStatus]})&"Ledger[Amount]"',
      );
      expect(files.columnConfigs).toContain('"formula": "=${amount}*2"');
    });

    it("records the column type, with the rule a dropdown carries", () => {
      seedValueSources({
        apiTables: [
          listSheet([{ header: "Payment status", values: ["Paid"] }]),
          ledgerSheet([
            { header: "Amount", values: [10], columnType: "CURRENCY" },
            {
              header: "Status",
              values: ["Paid"],
              columnType: "DROPDOWN",
              validationConditionType: "ONE_OF_RANGE",
              validationValues: ["=List[Payment status]"],
            },
            { header: "Note", values: ["x"] },
          ]),
        ],
      });

      const { ledger } = generateLedgerColumns();

      expect(ledger.amount?.columnType).toEqual({
        type: "CURRENCY",
        rule: null,
      });
      expect(ledger.status?.columnType).toEqual({
        type: "DROPDOWN",
        rule: { condition: "ONE_OF_RANGE", values: ["=List[Payment status]"] },
      });
      expect(ledger.note?.columnType).toEqual({ type: null, rule: null });
    });

    it("records each action-row cell as empty, a checkbox or a level-2 group heading", () => {
      seedValueSources({
        apiTables: [
          ledgerSheet([
            { header: "Amount", values: [10] },
            { header: "Run", values: [false], actionRowCell: false },
            { header: "Due", values: [1], actionRowCell: "Dates" },
          ]),
        ],
      });

      const { ledger } = generateLedgerColumns();

      expect(ledger.amount?.actionRowCell).toBeNull();
      expect(ledger.run?.actionRowCell).toEqual({ kind: "checkbox" });
      expect(ledger.due?.actionRowCell).toEqual({
        kind: "groupHeading2",
        text: "Dates",
      });
    });

    it("records a spanning level-1 group heading on its first cell only", () => {
      seedValueSources({
        apiTables: [
          ledgerSheet([
            { header: "Amount", values: [10], groupHeading1: "Money" },
            { header: "Fee", values: [1] },
          ]),
        ],
      });

      const { ledger } = generateLedgerColumns();

      expect(ledger.amount?.groupHeading1).toBe("Money");
      expect(ledger.fee?.groupHeading1).toBeNull();
    });
  });
});
