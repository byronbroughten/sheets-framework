import { beforeEach, describe, expect, it } from "vitest";

import { installedConfigs } from "../01_SpreadsheetSchema/configRegister";
import { configSheetFloorSeed } from "../01_SpreadsheetSchema/configSheetFloorSeed";
import { getTableTraitByName } from "../01_SpreadsheetSchema/tableConfigsTypes";
import { expectedSheetLayout } from "../testSupport/expectedSheetLayout";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  type FakeCellValue,
  type FakeSheetProperties,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { tableIdOnTab } from "../testSupport/fakeTableConfigSheet";
import {
  ConfigCoordinator,
  type ConfigRegeneration,
} from "./ConfigCoordinator";

const { columnConfigs } = installedConfigs();
const testSheetGid = getTableTraitByName("item", "sheetGid");
const tableConfigGid = 210603630;
const columnConfigGid = 2034522667;
const draftGid = 777000111;
const draftTitle = "Add Widget Order";
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

// The columns tableLayout replaced, as a spreadsheet synced before it still has them.
const legacyLayoutColumns = [
  { columnId: "c:sscf:XOpXA8U", header: "ID header", value: "ID" },
  { columnId: "c:sscf:Gp3PuNE", header: "Name header", value: "Name" },
  { columnId: "c:sscf:8uxVA53", header: "ID delimiter", value: ":" },
  {
    columnId: "c:sscf:RtBaCIb",
    header: "Start table column index base 1",
    value: 1,
  },
  {
    columnId: "c:sscf:Kt9oKSY",
    header: "Column ID row index base 1",
    value: 1,
  },
  {
    columnId: "c:sscf:Tm9zOUP",
    header: "Column group heading row index base 1",
    value: 2,
  },
  { columnId: "c:sscf:GKSJHu0", header: "Action row index base 1", value: 3 },
  {
    columnId: "c:sscf:58r8zkF",
    header: "Table header row index base 1",
    value: 4,
  },
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

function spreadsheetConfigSheet(
  options: {
    tableEndRowIndex?: number;
    fillRowIdsRunStatusColumnId?: string;
    hasLegacyLayoutColumns?: boolean;
    extraRows?: Record<number, readonly (string | number | boolean | null)[]>;
    protectedRanges?: GoogleAppsScript.Sheets.Schema.ProtectedRange[];
  } = {},
) {
  const legacyColumns = options.hasLegacyLayoutColumns
    ? legacyLayoutColumns
    : [];
  const headers = [
    ...sscColumns.map((columnName) => ssc[columnName].header),
    ...legacyColumns.map((column) => column.header),
  ];
  const groupHeadings = headers.map((header) =>
    spreadsheetConfigGroupHeading(header),
  );
  const dataRow = [
    ...sscColumns.map((columnName) =>
      columnName === "tableMenuSpace" ? "Not used" : "",
    ),
    ...legacyColumns.map((column) => column.value),
  ];
  return {
    sheetId: spreadsheetConfigGid,
    title: "Spreadsheet Config",
    rows: buildGridRows({
      0: [
        ...sscColumns.map((columnName) =>
          columnName === "fillRowIdsRunStatus"
            ? (options.fillRowIdsRunStatusColumnId ?? ssc[columnName].columnId)
            : ssc[columnName].columnId,
        ),
        ...legacyColumns.map((column) => column.columnId),
      ],
      1: groupHeadings,
      3: headers,
      4: dataRow,
      ...options.extraRows,
    }),
    table: {
      tableId: tableIdOnTab(spreadsheetConfigGid),
      name: configSheetFloorSeed.spreadsheetConfig.liveTableName,
      startRowIndex: tableHeaderRowIndex,
      startColumnIndex: startTableColIndex,
      endRowIndex: options.tableEndRowIndex ?? 5,
      endColumnIndex: headers.length,
      columnTypes: columnTypesByHeader("spreadsheetConfig", headers),
    },
    protectedRanges: options.protectedRanges,
  };
}

function headerOnlyDraftSheet(
  options: {
    sheetId?: number;
    title?: string;
    table?: "header-only" | "with-data-row" | "missing";
    hasIdHeader?: boolean;
  } = {},
): FakeSheetProperties {
  const table = options.table ?? "header-only";
  const headers = options.hasIdHeader ? ["ID", "Name"] : ["Name"];
  const rows: Record<number, readonly (string | number | boolean | null)[]> = {
    3: headers,
  };
  if (table === "with-data-row") {
    rows[4] = [];
  }
  const sheet: FakeSheetProperties = {
    sheetId: options.sheetId ?? draftGid,
    title: options.title ?? draftTitle,
    rows: buildGridRows(rows),
  };
  if (table === "missing") {
    return sheet;
  }
  return {
    ...sheet,
    table: {
      name: draftTitle,
      endRowIndex: table === "header-only" ? headerOnlyTableEndRowIndex : 5,
    },
  };
}

function seedFixture(
  options: {
    testColumnId?: string;
    testTableId?: string;
    itemHeaderRowIndex?: number;
    fillRowIdsRunStatusColumnId?: string;
    spreadsheetConfigTableEndRowIndex?: number;
    spreadsheetConfigExtraRows?: Record<
      number,
      readonly (string | number | boolean | null)[]
    >;
    valueConfigSheet?: FakeSheetProperties;
    extraSheets?: FakeSheetProperties[];
    extraTableConfigDataRows?: Record<
      number,
      readonly (string | number | boolean | null)[]
    >;
    tableConfigTableEndRowIndex?: number;
    hasLegacyLayoutColumns?: boolean;
    columnConfigDataRows?: readonly (readonly FakeCellValue[])[];
    spreadsheetConfigProtections?: GoogleAppsScript.Sheets.Schema.ProtectedRange[];
    isDryRun?: boolean;
  } = {},
) {
  const testColumnId = options.testColumnId ?? "c:itm:xyz123";
  const itemHeaderRowIndex = options.itemHeaderRowIndex ?? tableHeaderRowIndex;
  const columnConfigDataRows = options.columnConfigDataRows ?? [];
  return stubSheetsService({
    sheets: [
      spreadsheetConfigSheet({
        tableEndRowIndex: options.spreadsheetConfigTableEndRowIndex,
        fillRowIdsRunStatusColumnId: options.fillRowIdsRunStatusColumnId,
        hasLegacyLayoutColumns: options.hasLegacyLayoutColumns,
        extraRows: options.spreadsheetConfigExtraRows,
        protectedRanges: options.spreadsheetConfigProtections,
      }),
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
          4: [options.testTableId ?? "item", "", "Item", true],
          ...options.extraTableConfigDataRows,
        }),
        table: {
          tableId: tableIdOnTab(tableConfigGid),
          name: configSheetFloorSeed.tableConfig.liveTableName,
          startRowIndex: tableHeaderRowIndex,
          startColumnIndex: startTableColIndex,
          endRowIndex: options.tableConfigTableEndRowIndex ?? 5,
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
        // appendRowWithVals (used by ColumnConfigOperator._appendColumnRows)
        // resolves every non-formula column of the sheet's schema against
        // this row, not just the ones toFileSource reads — so all of the
        // real committed Column Config columns need to be present here.
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
          ...rowsFromFirstDataRow(columnConfigDataRows),
        }),
        table: {
          tableId: tableIdOnTab(columnConfigGid),
          name: configSheetFloorSeed.columnConfig.liveTableName,
          startRowIndex: tableHeaderRowIndex,
          startColumnIndex: startTableColIndex,
          endRowIndex:
            headerOnlyTableEndRowIndex +
            Math.max(1, columnConfigDataRows.length),
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
          tableId: options.testTableId ?? "item",
          name: "Item",
          startRowIndex: itemHeaderRowIndex,
          endRowIndex: itemHeaderRowIndex + 2,
        },
      },
      options.valueConfigSheet ?? floorValueConfigTab(),
      ...(options.extraSheets ?? []),
    ],
    isDryRun: options.isDryRun,
  });
}

const valueConfigFloorGid = getTableTraitByName("valueConfig", "sheetGid");

function floorValueConfigTab(
  tableId = tableIdOnTab(valueConfigFloorGid),
): FakeSheetProperties {
  return {
    ...valueConfigTab({
      sheetId: valueConfigFloorGid,
      title: "Value Config",
      columnId: "c:vcf:abc1234",
    }),
    table: {
      tableId,
      name: configSheetFloorSeed.valueConfig.liveTableName,
      startRowIndex: tableHeaderRowIndex,
      startColumnIndex: startTableColIndex,
      endRowIndex: 5,
      endColumnIndex: startTableColIndex + 1,
      columnTypes: { 0: "TEXT" },
    },
  };
}

function valueConfigTab(options: {
  sheetId: number;
  title: string;
  columnId?: string;
}): FakeSheetProperties {
  return {
    sheetId: options.sheetId,
    title: options.title,
    rows: buildGridRows({
      0: [options.columnId ?? ""],
      3: ["Value title"],
      4: [],
    }),
    table: {
      tableId: tableIdOnTab(options.sheetId),
      name: configSheetFloorSeed.valueConfig.liveTableName,
      endRowIndex: 5,
    },
  };
}

function rowsFromFirstDataRow(
  rows: readonly (readonly FakeCellValue[])[],
): Record<number, readonly FakeCellValue[]> {
  return rows.reduce<Record<number, readonly FakeCellValue[]>>(
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

function driftedFloorWarningDescription(): string {
  return "Config-sheet floor · Spreadsheet Config · warning";
}

function tableConfigLetApiAccess(
  orchestrator: ConfigCoordinator,
  sheetGid: number,
): boolean | "" {
  const sheet = orchestrator.tableConfigOperator.table;
  const col = sheet.columns("tableId", "letApiAccess");
  const rowIndex = sheet.workingRowIndexesWithData.find(
    (index) => col.tableId.value(index) === tableIdOnTab(sheetGid),
  );
  if (rowIndex === undefined) {
    throw new Error(`No Table Config row for gid ${sheetGid}`);
  }
  return col.letApiAccess.valueOrEmpty(rowIndex);
}

interface TableConfigsEntryFixture {
  tableKey: string;
  tableName: string;
  sheetGid: number;
  idPrefix: string;
  hasIdColumn: boolean;
}

// Every fixture Table sits at the expected origin and carries a "Name" header.
function tableConfigsEntry({
  tableKey,
  tableName,
  sheetGid,
  idPrefix,
  hasIdColumn,
}: TableConfigsEntryFixture): string {
  return `"${tableKey}": { "tableId": "${tableIdOnTab(sheetGid)}", "tableName": "${tableName}", "sheetGid": ${sheetGid}, "idPrefix": "${idPrefix}", "hasIdColumn": ${hasIdColumn}, "hasNameColumn": true }`;
}

function driftedFloorWarningProtection(): GoogleAppsScript.Sheets.Schema.ProtectedRange {
  return {
    protectedRangeId: 41,
    description: driftedFloorWarningDescription(),
    warningOnly: true,
    range: { sheetId: spreadsheetConfigGid },
    unprotectedRanges: [
      {
        sheetId: spreadsheetConfigGid,
        startRowIndex: 0,
        endRowIndex: 1,
        startColumnIndex: 0,
        endColumnIndex: 1,
      },
    ],
  };
}

describe("ConfigCoordinator.syncAndFlushConfigSheets", () => {
  it("flushes the config-sheet floor and the config sheets' changes in two batch updates", () => {
    const { batchUpdateCount, grid } = seedFixture();

    const orchestrator = ConfigCoordinator.init();
    orchestrator.syncAndFlushConfigSheets();

    expect(batchUpdateCount()).toBe(2);
    expect(
      grid
        .sheet(spreadsheetConfigGid)
        .protectedRanges.map((protection) => protection.description),
    ).toContain(driftedFloorWarningDescription());
    expect(
      grid.sheet(columnConfigGid).values({
        startRowIndex: 4,
        endRowIndex: 5,
        endColumnIndex: 4,
      }),
    ).toEqual([["item", "c:itm:xyz123", "Item", "Some Header"]]);
    // The gathered column ID rode the appended Column Config row into the flush.
    expect(orchestrator.tableConfigOperator.newTableConfigs().item).toEqual({
      tableId: "item",
      tableName: "Item",
      sheetGid: testSheetGid,
      idPrefix: "itm",
      hasIdColumn: false,
      hasNameColumn: false,
    });
  });

  it("returns the untyped-column summary for the endpoint to report, with no file sources", () => {
    seedFixture();

    const summary = ConfigCoordinator.init().syncConfigSheetRows();

    expect(summary).toContain("1 column(s) across 1 Table(s)");
    expect(typeof summary).toBe("string");
  });

  it("returns the floor report beside the untyped-column summary, as one line", () => {
    seedFixture({
      spreadsheetConfigProtections: [driftedFloorWarningProtection()],
    });

    const summary = ConfigCoordinator.init().syncConfigSheetRows();

    expect(summary).toContain("Replaced drifted:");
    expect(summary).toContain(driftedFloorWarningDescription());
    expect(summary).toContain("1 column(s) across 1 Table(s)");
    expect(summary).not.toContain("\n");
  });
});

describe("ConfigCoordinator.generateConfigFiles", () => {
  it("returns every file's source together, reflecting the synced state", () => {
    seedFixture();

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
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

  it("emits no Spreadsheet Config file", () => {
    seedFixture();

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    expect(Object.keys(parsed)).not.toContain("spreadsheetConfig");
    expect(Object.values(parsed).join("\n")).not.toContain(
      "makeSpreadsheetConfig",
    );
  });

  it("syncs a Spreadsheet Config that still has the old layout columns, cataloguing them as plain columns", () => {
    seedFixture({ hasLegacyLayoutColumns: true });

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    legacyLayoutColumns.forEach(({ columnId }) => {
      expect(parsed.columnConfigs).toContain(columnId);
    });
  });

  it("ignores edited old layout values, since nothing reads them", () => {
    seedFixture({
      hasLegacyLayoutColumns: true,
      spreadsheetConfigExtraRows: {
        4: ["Not used", "", "", "", "", "Key", "Title", "|", 9, 9, 9, 9, 9],
      },
    });

    expect(() =>
      ConfigCoordinator.init().syncAndFlushConfigSheets(),
    ).not.toThrow();
  });

  it("prunes the old layout columns' Column Config rows once they are gone from the sheet", () => {
    const legacyColumnConfigRows = legacyLayoutColumns.map((column) => [
      tableIdOnTab(spreadsheetConfigGid),
      column.columnId,
      "Spreadsheet Config",
      column.header,
      false,
    ]);
    seedFixture({ columnConfigDataRows: legacyColumnConfigRows });

    const coordinator = ConfigCoordinator.init();
    const parsed = coordinator.generateConfigFiles("../makeConfigs");
    const columnIdColumn =
      coordinator.columnConfigOperator.table.column("columnId");
    legacyLayoutColumns.forEach(({ columnId }) => {
      expect(columnIdColumn.hasValue(columnId)).toBe(false);
      expect(parsed.columnConfigs).not.toContain(columnId);
    });
  });

  it("carries the untyped-column summary back, since no run status cell will show it", () => {
    seedFixture();

    expect(
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs")
        .untypedColumnsSummary,
    ).toContain("1 column(s) across 1 Table(s)");
  });

  it("carries the floor report back beside the untyped-column summary", () => {
    seedFixture({
      spreadsheetConfigProtections: [driftedFloorWarningProtection()],
    });

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    expect(parsed.floorReport).toBe(
      `Replaced drifted: ${driftedFloorWarningDescription()}`,
    );
    expect(parsed.untypedColumnsSummary).toContain(
      "1 column(s) across 1 Table(s)",
    );
  });

  it("carries the declared-cell report back, since no run status cell will show it", () => {
    seedFixture({
      extraTableConfigDataRows: {
        5: [
          tableIdOnTab(spreadsheetConfigGid),
          "",
          "Spreadsheet Config",
          false,
          "",
        ],
      },
      tableConfigTableEndRowIndex: 6,
    });

    expect(
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs")
        .declaredCellReport,
    ).toContain("Table Config · Let api access · Spreadsheet Config → TRUE");
  });

  it("still catalogs value titles after Column Config pruned a stale row of its own", () => {
    const columnIdRow = [
      cc.tableId.columnId,
      cc.columnId.columnId,
      cc.tableName.columnId,
      cc.header.columnId,
      cc.emptyValueAllowed.columnId,
    ];
    stubSheetsService({
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
            4: [tableIdOnTab(columnConfigGid), "", "Column Config", true],
          }),
          table: {
            tableId: tableIdOnTab(tableConfigGid),
            name: configSheetFloorSeed.tableConfig.liveTableName,
            startRowIndex: tableHeaderRowIndex,
            startColumnIndex: startTableColIndex,
            endRowIndex: 5,
          },
        },
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnIdRow,
            3: [
              cc.tableId.header,
              cc.columnId.header,
              cc.tableName.header,
              cc.header.header,
              cc.emptyValueAllowed.header,
            ],
            4: [
              tableIdOnTab(columnConfigGid),
              cc.tableId.columnId,
              "Column Config",
              cc.tableId.header,
            ],
            5: [
              tableIdOnTab(columnConfigGid),
              "c:ccf:stale-gone",
              "Column Config",
              "Gone",
            ],
          }),
          table: {
            tableId: tableIdOnTab(columnConfigGid),
            name: configSheetFloorSeed.columnConfig.liveTableName,
            startRowIndex: tableHeaderRowIndex,
            startColumnIndex: startTableColIndex,
            endRowIndex: 6,
          },
        },
        floorValueConfigTab(),
      ],
    });

    expect(() =>
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs"),
    ).not.toThrow();
  });

  // Dry runs, since once the floor and sync writes land the guard never sees the change (#15).
  describe("floor identity", () => {
    const movedValueConfigGid = 999000111;

    function seedValueConfigTab(tab: FakeSheetProperties) {
      seedFixture({
        valueConfigSheet: tab,
        extraTableConfigDataRows: {
          5: [tableIdOnTab(tab.sheetId), "", tab.title, true, ""],
        },
        tableConfigTableEndRowIndex: 6,
        isDryRun: true,
      });
    }

    it("throws naming the floor tab, its previous GID and its new one, and returns no file source", () => {
      seedValueConfigTab(
        valueConfigTab({ sheetId: movedValueConfigGid, title: "Value config" }),
      );

      expect(() =>
        ConfigCoordinator.init().generateConfigFiles("../makeConfigs"),
      ).toThrow(
        `Floor tab "valueConfig" GID was ${valueConfigFloorGid} and is now ${movedValueConfigGid}.`,
      );
    });

    it("throws when a floor tab's ID prefix changed", () => {
      seedValueConfigTab(
        valueConfigTab({
          sheetId: valueConfigFloorGid,
          title: "Value Config",
          columnId: "c:zzz:abc1234",
        }),
      );

      expect(() =>
        ConfigCoordinator.init().generateConfigFiles("../makeConfigs"),
      ).toThrow(
        'Floor tab "valueConfig" ID prefix was "vcf" and is now "zzz".',
      );
    });

    it("throws when a floor tab's Table ID changed", () => {
      const replacedTableId = "tbl-replaced01";
      seedFixture({
        valueConfigSheet: floorValueConfigTab(replacedTableId),
        extraTableConfigDataRows: {
          5: [replacedTableId, "", "Value Config", true, ""],
        },
        tableConfigTableEndRowIndex: 6,
        isDryRun: true,
      });

      expect(() =>
        ConfigCoordinator.init().generateConfigFiles("../makeConfigs"),
      ).toThrow(
        `Floor tab "valueConfig" Table ID was "${tableIdOnTab(valueConfigFloorGid)}" and is now "${replacedTableId}".`,
      );
    });

    it("fails a floor column's changed column ID with the identity guard's message, before the floor seed check", () => {
      seedFixture({
        fillRowIdsRunStatusColumnId: "c:sscf:moved01",
        isDryRun: true,
      });

      expect(() =>
        ConfigCoordinator.init().generateConfigFiles("../makeConfigs"),
      ).toThrow(
        `Floor column "Fill row IDs, run status" on "spreadsheetConfig" had column ID "${ssc.fillRowIdsRunStatus.columnId}" and is now "c:sscf:moved01".`,
      );
    });

    it("still only reports a non-floor sheet's changed ID prefix", () => {
      seedFixture({ testColumnId: "c:zzz:xyz123" });

      expect(
        ConfigCoordinator.init().generateConfigFiles("../makeConfigs")
          .idPrefixReport,
      ).toBe(
        'Table "Item" sampled ID prefix "zzz" differs from last generated "itm".',
      );
    });
  });
});

describe("ConfigCoordinator.syncConfigSheetRows Spreadsheet Config Table", () => {
  it("proceeds when Spreadsheet Config's Table has exactly one data row", () => {
    seedFixture();

    expect(() => ConfigCoordinator.init().syncConfigSheetRows()).not.toThrow();
  });

  it("throws when Spreadsheet Config's Table has two data rows", () => {
    seedFixture({ spreadsheetConfigTableEndRowIndex: 6 });

    expect(() => ConfigCoordinator.init().syncConfigSheetRows()).toThrow(
      /Spreadsheet Config/,
    );
  });

  it("throws when Spreadsheet Config's Table has no data row", () => {
    seedFixture({ spreadsheetConfigTableEndRowIndex: 4 });

    expect(() => ConfigCoordinator.init().syncConfigSheetRows()).toThrow(
      /Spreadsheet Config/,
    );
  });

  it("ignores a filled row below Spreadsheet Config's Table", () => {
    seedFixture({
      spreadsheetConfigExtraRows: {
        5: ["junk", "", "", "", "", "", ""],
      },
    });

    expect(() => ConfigCoordinator.init().syncConfigSheetRows()).not.toThrow();
  });
});

describe("ConfigCoordinator.generateConfigFiles Spreadsheet Config Table", () => {
  it("throws when Spreadsheet Config's Table has two data rows", () => {
    seedFixture({ spreadsheetConfigTableEndRowIndex: 6 });

    expect(() =>
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs"),
    ).toThrow(/Spreadsheet Config/);
  });
});

describe("ConfigCoordinator.syncConfigSheetRows Let api access", () => {
  it("syncs when Let api access is off and the Table has no data row, cataloguing the tab without emitting it", () => {
    seedFixture({
      extraSheets: [headerOnlyDraftSheet()],
      extraTableConfigDataRows: {
        5: [tableIdOnTab(draftGid), "", draftTitle, false, ""],
      },
      tableConfigTableEndRowIndex: 6,
    });

    const orchestrator = ConfigCoordinator.init();
    expect(() => orchestrator.syncConfigSheetRows()).not.toThrow();
    expect(
      orchestrator.tableConfigOperator.newTableConfigs().addWidgetOrder,
    ).toBeUndefined();
    expect(
      orchestrator.tableConfigOperator.table
        .column("tableId")
        .hasValue(tableIdOnTab(draftGid)),
    ).toBe(true);
  });

  it("aborts and names the tab when Let api access is on and the Table has no data row", () => {
    seedFixture({
      extraSheets: [headerOnlyDraftSheet()],
      extraTableConfigDataRows: {
        5: [tableIdOnTab(draftGid), "", draftTitle, true, ""],
      },
      tableConfigTableEndRowIndex: 6,
    });

    expect(() => ConfigCoordinator.init().syncConfigSheetRows()).toThrow(
      /Add Widget Order.*has only its header/,
    );
  });

  it("samples hasIdColumn from the Table header row of a Let api access sheet", () => {
    seedFixture({
      extraSheets: [
        headerOnlyDraftSheet({
          table: "with-data-row",
          hasIdHeader: true,
        }),
      ],
      extraTableConfigDataRows: {
        5: [tableIdOnTab(draftGid), "", draftTitle, true, "awo"],
      },
      tableConfigTableEndRowIndex: 6,
    });

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    expect(parsed.tableConfigs).toContain(
      tableConfigsEntry({
        tableKey: "addWidgetOrder",
        tableName: draftTitle,
        sheetGid: draftGid,
        idPrefix: "awo",
        hasIdColumn: true,
      }),
    );
  });

  it("treats a never-ticked Let api access box as off", () => {
    seedFixture({
      extraSheets: [headerOnlyDraftSheet()],
      extraTableConfigDataRows: {
        5: [tableIdOnTab(draftGid), "", draftTitle, null, ""],
      },
      tableConfigTableEndRowIndex: 6,
    });

    const orchestrator = ConfigCoordinator.init();
    expect(() => orchestrator.syncConfigSheetRows()).not.toThrow();
    expect(
      orchestrator.tableConfigOperator.newTableConfigs().addWidgetOrder,
    ).toBeUndefined();
  });

  it("catalogues a brand-new header-only tab with Let api access off", () => {
    seedFixture({ extraSheets: [headerOnlyDraftSheet()] });

    const orchestrator = ConfigCoordinator.init();
    const parsed = orchestrator.generateConfigFiles("../makeConfigs");
    expect(
      orchestrator.tableConfigOperator.table
        .column("tableId")
        .hasValue(tableIdOnTab(draftGid)),
    ).toBe(true);
    expect(parsed.tableConfigs).not.toContain("addWidgetOrder");
  });

  it("omits a draft that has data rows until Let api access is ticked", () => {
    seedFixture({
      extraSheets: [headerOnlyDraftSheet({ table: "with-data-row" })],
      extraTableConfigDataRows: {
        5: [tableIdOnTab(draftGid), "", draftTitle, false, "awo"],
      },
      tableConfigTableEndRowIndex: 6,
    });

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    expect(parsed.tableConfigs).toContain('"item"');
    expect(parsed.tableConfigs).not.toContain("addWidgetOrder");
    expect(parsed.columnConfigs).toContain("c:itm:xyz123");
    expect(parsed.columnConfigs).not.toMatch(/c:awo:/);
  });

  it("syncs several header-only drafts in one run", () => {
    const secondDraftGid = draftGid + 1;
    const secondTitle = "Add Occ Charge Intention";
    seedFixture({
      extraSheets: [
        headerOnlyDraftSheet(),
        headerOnlyDraftSheet({
          sheetId: secondDraftGid,
          title: secondTitle,
        }),
      ],
    });

    const orchestrator = ConfigCoordinator.init();
    expect(() => orchestrator.syncConfigSheetRows()).not.toThrow();
    const tableIds = orchestrator.tableConfigOperator.table.column("tableId");
    expect(tableIds.hasValue(tableIdOnTab(draftGid))).toBe(true);
    expect(tableIds.hasValue(tableIdOnTab(secondDraftGid))).toBe(true);
  });

  it("regens a healthy Let api access sheet while cataloguing an empty draft", () => {
    seedFixture({ extraSheets: [headerOnlyDraftSheet()] });

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    expect(parsed.tableConfigs).toContain('"item"');
    expect(parsed.tableConfigs).not.toContain("addWidgetOrder");
    expect(parsed.columnConfigs).toContain("c:itm:xyz123");
  });

  it("still syncs the config-describing sheets when they have Let api access", () => {
    seedFixture({
      extraTableConfigDataRows: {
        5: [tableIdOnTab(tableConfigGid), "", "Table Config", true, "scf"],
        6: [tableIdOnTab(columnConfigGid), "", "Column Config", true, "ccf"],
      },
      tableConfigTableEndRowIndex: 7,
    });

    expect(() => ConfigCoordinator.init().syncConfigSheetRows()).not.toThrow();
  });

  it("writes an unticked Let api access on a floor tab back to TRUE and names the tab", () => {
    seedFixture({
      extraTableConfigDataRows: {
        5: [
          tableIdOnTab(spreadsheetConfigGid),
          "",
          "Spreadsheet Config",
          false,
          "",
        ],
      },
      tableConfigTableEndRowIndex: 6,
    });

    const orchestrator = ConfigCoordinator.init();
    const report = orchestrator.syncConfigSheetRows();

    expect(tableConfigLetApiAccess(orchestrator, spreadsheetConfigGid)).toBe(
      true,
    );
    expect(report).toContain(
      "Table Config · Let api access · Spreadsheet Config → TRUE",
    );
  });

  it("ticks all four floor Tables' unticked Let api access rows again in one sync", () => {
    const floorGids = [
      spreadsheetConfigGid,
      tableConfigGid,
      columnConfigGid,
      valueConfigFloorGid,
    ];
    seedFixture({
      extraTableConfigDataRows: Object.fromEntries(
        floorGids.map((sheetGid, index) => [
          5 + index,
          [tableIdOnTab(sheetGid), "", "", false, ""],
        ]),
      ),
      tableConfigTableEndRowIndex: 5 + floorGids.length,
    });

    const orchestrator = ConfigCoordinator.init();
    orchestrator.syncConfigSheetRows();

    expect(
      floorGids.map((sheetGid) =>
        tableConfigLetApiAccess(orchestrator, sheetGid),
      ),
    ).toEqual([true, true, true, true]);
  });

  it("appends a missing floor-tab row with Let api access TRUE", () => {
    seedFixture();

    const orchestrator = ConfigCoordinator.init();
    const report = orchestrator.syncConfigSheetRows();

    expect(
      orchestrator.tableConfigOperator.table
        .column("tableId")
        .hasValue(tableIdOnTab(spreadsheetConfigGid)),
    ).toBe(true);
    expect(tableConfigLetApiAccess(orchestrator, spreadsheetConfigGid)).toBe(
      true,
    );
    expect(report).toContain(
      "Table Config · Let api access · Spreadsheet Config → TRUE",
    );
  });

  it("leaves a non-floor tab's unticked Let api access alone", () => {
    seedFixture({
      extraSheets: [headerOnlyDraftSheet({ table: "with-data-row" })],
      extraTableConfigDataRows: {
        5: [tableIdOnTab(draftGid), "", draftTitle, false, "awo"],
      },
      tableConfigTableEndRowIndex: 6,
    });

    const orchestrator = ConfigCoordinator.init();
    const report = orchestrator.syncConfigSheetRows();

    expect(tableConfigLetApiAccess(orchestrator, draftGid)).toBe(false);
    expect(report ?? "").not.toContain(draftTitle);
  });

  it("drops a ticked row whose Table is gone, since a tab with no Table gets no row", () => {
    seedFixture({
      extraSheets: [headerOnlyDraftSheet({ table: "missing" })],
      extraTableConfigDataRows: {
        5: [tableIdOnTab(draftGid), "", draftTitle, true, "awo"],
      },
      tableConfigTableEndRowIndex: 6,
    });

    const orchestrator = ConfigCoordinator.init();
    expect(() => orchestrator.syncConfigSheetRows()).not.toThrow();
    expect(
      orchestrator.tableConfigOperator.table
        .column("sheetTitle")
        .hasValue(draftTitle),
    ).toBe(false);
  });

  it("syncs when a draft tab has no Table", () => {
    seedFixture({
      extraSheets: [headerOnlyDraftSheet({ table: "missing" })],
    });

    expect(() => ConfigCoordinator.init().syncConfigSheetRows()).not.toThrow();
  });
});

describe("ConfigCoordinator.generateConfigFiles ID prefix", () => {
  const widgetGid = 888001;
  const otherGid = 888002;

  function letApiAccessSheet(options: {
    sheetId: number;
    title: string;
    columnIds?: readonly string[];
    headers?: readonly string[];
  }): FakeSheetProperties {
    const headers = options.headers ?? ["Name"];
    const columnIds = options.columnIds ?? headers.map(() => "");
    return {
      sheetId: options.sheetId,
      title: options.title,
      rows: buildGridRows({
        0: columnIds,
        3: headers,
        4: [],
      }),
      table: { name: options.title, endRowIndex: 5 },
    };
  }

  function tableConfigRow(props: {
    sheetId: number;
    title: string;
    leftoverIdPrefix: string;
  }) {
    return [
      tableIdOnTab(props.sheetId),
      "",
      props.title,
      true,
      props.leftoverIdPrefix,
    ];
  }

  it("assigns a generated prefix to a new Let api access sheet and mints column IDs with it", () => {
    seedFixture({
      extraSheets: [letApiAccessSheet({ sheetId: widgetGid, title: "Widget" })],
      extraTableConfigDataRows: {
        5: tableConfigRow({
          sheetId: widgetGid,
          title: "Widget",
          leftoverIdPrefix: "zzzz",
        }),
      },
      tableConfigTableEndRowIndex: 6,
    });

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    expect(parsed.tableConfigs).toContain(
      tableConfigsEntry({
        tableKey: "widget",
        tableName: "Widget",
        sheetGid: widgetGid,
        idPrefix: "wdg",
        hasIdColumn: false,
      }),
    );
    expect(parsed.columnConfigs).toMatch(/c:wdg:/);
  });

  it("samples an existing sheet's prefix from its column IDs, whatever the tab title is now", () => {
    seedFixture({
      extraSheets: [
        letApiAccessSheet({
          sheetId: widgetGid,
          title: "Renamed Tab",
          columnIds: ["c:wdg:abc1234"],
        }),
      ],
      extraTableConfigDataRows: {
        5: tableConfigRow({
          sheetId: widgetGid,
          title: "Renamed Tab",
          leftoverIdPrefix: "zzzz",
        }),
      },
      tableConfigTableEndRowIndex: 6,
    });

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    expect(parsed.tableConfigs).toContain(
      tableConfigsEntry({
        tableKey: "renamedTab",
        tableName: "Renamed Tab",
        sheetGid: widgetGid,
        idPrefix: "wdg",
        hasIdColumn: false,
      }),
    );
  });

  it("mints blank column ID cells on an existing sheet with the sampled prefix", () => {
    seedFixture({
      extraSheets: [
        letApiAccessSheet({
          sheetId: widgetGid,
          title: "Widget",
          columnIds: ["c:wdg:abc1234", ""],
          headers: ["Name", "Notes"],
        }),
      ],
      extraTableConfigDataRows: {
        5: tableConfigRow({
          sheetId: widgetGid,
          title: "Widget",
          leftoverIdPrefix: "zzzz",
        }),
      },
      tableConfigTableEndRowIndex: 6,
    });

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    const minted = parsed.columnConfigs.match(/c:wdg:[^"]+/g) ?? [];
    expect(minted).toHaveLength(2);
    expect(minted).toContain("c:wdg:abc1234");
  });

  it("stops when one sheet's column IDs carry mixed prefixes, naming the sheet and stray IDs", () => {
    seedFixture({
      extraSheets: [
        letApiAccessSheet({
          sheetId: widgetGid,
          title: "Widget",
          columnIds: ["c:wdg:abc1234", "c:gzm:xyz1234"],
          headers: ["Name", "Notes"],
        }),
      ],
      extraTableConfigDataRows: {
        5: tableConfigRow({
          sheetId: widgetGid,
          title: "Widget",
          leftoverIdPrefix: "wdg",
        }),
      },
      tableConfigTableEndRowIndex: 6,
    });

    expect(() => ConfigCoordinator.init().syncConfigSheetRows()).toThrow(
      /Widget.*c:gzm:xyz1234/,
    );
  });

  it("stops when two sheets share a sampled prefix, named by sheet title", () => {
    seedFixture({
      extraSheets: [
        letApiAccessSheet({
          sheetId: widgetGid,
          title: "Widget",
          columnIds: ["c:wdg:aaa1111"],
        }),
        letApiAccessSheet({
          sheetId: otherGid,
          title: "Gizmo",
          columnIds: ["c:wdg:bbb2222"],
        }),
      ],
      extraTableConfigDataRows: {
        5: tableConfigRow({
          sheetId: widgetGid,
          title: "Widget",
          leftoverIdPrefix: "wdg",
        }),
        6: tableConfigRow({
          sheetId: otherGid,
          title: "Gizmo",
          leftoverIdPrefix: "gzm",
        }),
      },
      tableConfigTableEndRowIndex: 7,
    });

    expect(() =>
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs"),
    ).toThrow(/Widget.*Gizmo.*"wdg"/);
  });

  it("reports a sampled prefix that differs from the last generated table configs without failing", () => {
    seedFixture({ testColumnId: "c:zzz:xyz123" });

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    expect(parsed.idPrefixReport).toBe(
      'Table "Item" sampled ID prefix "zzz" differs from last generated "itm".',
    );
    expect(parsed.tableConfigs).toContain('"idPrefix": "zzz"');
  });

  it("fails when two Tables' names give the same key, naming both", () => {
    seedFixture({
      extraSheets: [
        {
          ...letApiAccessSheet({ sheetId: widgetGid, title: "Widget" }),
          table: { name: "item", endRowIndex: 5 },
        },
      ],
      extraTableConfigDataRows: {
        5: [tableIdOnTab(widgetGid), "", "Widget", true, "wdg"],
      },
      tableConfigTableEndRowIndex: 6,
    });

    expect(() =>
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs"),
    ).toThrow('Tables "Item" and "item" both give the key "item".');
  });

  it("compares a Table with the previous entry by tableId, not by sheet GID", () => {
    seedFixture({
      testColumnId: "c:zzz:xyz123",
      testTableId: "item-recreated",
    });

    expect(
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs")
        .idPrefixReport,
    ).toBeUndefined();
  });

  it("gives a tab without Let api access no prefix and no column IDs", () => {
    seedFixture({
      extraSheets: [
        letApiAccessSheet({
          sheetId: widgetGid,
          title: "Widget",
          columnIds: [""],
        }),
      ],
      extraTableConfigDataRows: {
        5: [tableIdOnTab(widgetGid), "", "Widget", false, "wdg"],
      },
      tableConfigTableEndRowIndex: 6,
    });

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    expect(parsed.tableConfigs).not.toContain("widget");
    expect(parsed.columnConfigs).not.toMatch(/c:wdg:/);
  });

  it("avoids prefixes sampled from other Let api access sheets in the same run", () => {
    seedFixture({
      extraSheets: [
        letApiAccessSheet({
          sheetId: widgetGid,
          title: "Gadget",
          columnIds: ["c:wdg:abc1234"],
        }),
        letApiAccessSheet({
          sheetId: otherGid,
          title: "Widget",
        }),
      ],
      extraTableConfigDataRows: {
        5: tableConfigRow({
          sheetId: widgetGid,
          title: "Gadget",
          leftoverIdPrefix: "gdg",
        }),
        6: tableConfigRow({
          sheetId: otherGid,
          title: "Widget",
          leftoverIdPrefix: "zzzz",
        }),
      },
      tableConfigTableEndRowIndex: 7,
    });

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    expect(parsed.tableConfigs).toContain(
      tableConfigsEntry({
        tableKey: "gadget",
        tableName: "Gadget",
        sheetGid: widgetGid,
        idPrefix: "wdg",
        hasIdColumn: false,
      }),
    );
    expect(parsed.tableConfigs).toContain(
      tableConfigsEntry({
        tableKey: "widget",
        tableName: "Widget",
        sheetGid: otherGid,
        idPrefix: "wdgt",
        hasIdColumn: false,
      }),
    );
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

  function recordsTableId(table: RecordsTable): string {
    return `records-${table.name.toLowerCase()}`;
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
    const tableConfigRows = tables.reduce<
      Record<number, readonly FakeCellValue[]>
    >((byIndex, table, offset) => {
      byIndex[5 + offset] = [
        recordsTableId(table),
        "",
        "Records",
        table.isManaged,
      ];
      return byIndex;
    }, {});
    return seedFixture({
      extraSheets: [recordsSheet(tables)],
      extraTableConfigDataRows: tableConfigRows,
      tableConfigTableEndRowIndex: 5 + tables.length,
    });
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
      parsed = ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
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

  it("refuses a ticked Table stacked clear of another but below the header zone, naming it", () => {
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

    expect(() =>
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs"),
    ).toThrowError(
      /^Table "Tenants" on "Records" \(gid \d+\) must have its header row on row 4\.$/,
    );
  });

  it("refuses a recorded, ticked Table moved below the header zone with generation's own fix", () => {
    seedFixture({ itemHeaderRowIndex: tableHeaderRowIndex + 2 });

    expect(() =>
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs"),
    ).toThrowError(
      /^Table "Item" on "Item" \(gid \d+\) must have its header row on row 4\.$/,
    );
  });

  it("drops a recorded Table moved below the header zone once it is unticked", () => {
    seedFixture({
      itemHeaderRowIndex: tableHeaderRowIndex + 2,
      extraTableConfigDataRows: { 4: ["item", "", "Item", false] },
    });

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    expect(parsed.tableConfigs).not.toContain('"item"');
  });

  it("leaves an unticked Table below the header zone to the operator", () => {
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

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
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

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    expect(parsed.tableConfigs).toContain('"tenants"');
    expect(parsed.tableConfigs).toContain('"leases"');
  });
});
