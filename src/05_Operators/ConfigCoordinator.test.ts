import { afterEach, beforeEach, describe, expect, it } from "vitest";

import { installedConfigs } from "../01_SpreadsheetSchema/configRegister";
import { configSheetFloorSeed } from "../01_SpreadsheetSchema/configSheetFloorSeed";
import { getSheetTraitByName } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import {
  clearSpreadsheetConfigOverlay,
  ssConfigGet,
} from "../01_SpreadsheetSchema/spreadsheetConfigTypes";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  type FakeSheetProperties,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { ConfigCoordinator } from "./ConfigCoordinator";

const { columnConfigs, spreadsheetConfig } = installedConfigs();
const testSheetGid = getSheetTraitByName("item", "sheetGid");
const sheetConfigGid = 210603630;
const columnConfigGid = 2034522667;
const draftGid = 777000111;
const draftTitle = "Add Widget Order";
const headerOnlyTableEndRowIndex = ssConfigGet("tableHeaderRowIndexBase0") + 1;
const tableHeaderRowIndex = ssConfigGet("tableHeaderRowIndexBase0");
const startTableColIndex = ssConfigGet("startTableColIndexBase0");
const spreadsheetConfigGid = getSheetTraitByName(
  "spreadsheetConfig",
  "sheetGid",
);

const sc = columnConfigs.sheetConfig;
const cc = columnConfigs.columnConfig;
const ssc = columnConfigs.spreadsheetConfig;

const sscColumns = [
  "tableMenuSpace",
  "fillRowIdsTimeLastRan",
  "fillRowIdsRunStatus",
  "syncConfigSheetRowsTimeLastRan",
  "syncConfigSheetRowsRunStatus",
  "idDelimiter",
  "idHeader",
  "nameHeader",
  "startTableColumnIndexBase1",
  "columnIdRowIndexBase1",
  "columnGroupHeadingRowIndexBase1",
  "actionRowIndexBase1",
  "tableHeaderRowIndexBase1",
] as const;

function floorSeedType(
  sheetName: "spreadsheetConfig" | "sheetConfig" | "columnConfig",
  header: string,
): string | undefined {
  const column = configSheetFloorSeed[sheetName].columns.find(
    (entry) => entry.header === header,
  );
  if (column !== undefined) return column.columnType;
  if (sheetName !== "spreadsheetConfig") return undefined;
  return Object.values(configSheetFloorSeed.spreadsheetConfig.endpoints)
    .flatMap((endpoint) => [endpoint.timeLastRan, endpoint.runStatus])
    .find((entry) => entry.header === header)?.columnType;
}

function columnTypesByHeader(
  sheetName: "spreadsheetConfig" | "sheetConfig" | "columnConfig",
  headers: readonly string[],
): Record<number, string> {
  const types: Record<number, string> = {};
  headers.forEach((header, colIndex) => {
    const columnType = floorSeedType(sheetName, header);
    if (columnType !== undefined) types[colIndex] = columnType;
  });
  return types;
}

const testSheetConfigRowWithApiAccess = [testSheetGid, "Item", true];

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  clearSpreadsheetConfigOverlay();
});

function spreadsheetConfigSheet(
  idDelimiter: string,
  options: {
    idHeader?: string;
    tableEndRowIndex?: number;
    fillRowIdsRunStatusColumnId?: string;
    startTableColumnIndexBase1?: number;
    extraRows?: Record<number, readonly (string | number | boolean | null)[]>;
    protectedRanges?: GoogleAppsScript.Sheets.Schema.ProtectedRange[];
  } = {},
) {
  const headers = sscColumns.map((columnName) => ssc[columnName].header);
  const groupHeadings = sscColumns.map((columnName) =>
    spreadsheetConfigGroupHeading(ssc[columnName].header),
  );
  const dataRow = sscColumns.map((columnName) => {
    if (columnName === "tableMenuSpace") return "Not used";
    if (columnName === "idDelimiter") return idDelimiter;
    if (columnName === "idHeader") return options.idHeader ?? "ID";
    if (columnName === "nameHeader") return "Name";
    if (columnName === "startTableColumnIndexBase1") {
      return options.startTableColumnIndexBase1 ?? 1;
    }
    if (columnName === "columnIdRowIndexBase1") return 1;
    if (columnName === "columnGroupHeadingRowIndexBase1") return 2;
    if (columnName === "actionRowIndexBase1") return 3;
    if (columnName === "tableHeaderRowIndexBase1") return 4;
    return "";
  });
  return {
    sheetId: spreadsheetConfigGid,
    title: "Spreadsheet Config",
    rows: buildGridRows({
      0: sscColumns.map((columnName) =>
        columnName === "fillRowIdsRunStatus"
          ? (options.fillRowIdsRunStatusColumnId ?? ssc[columnName].columnId)
          : ssc[columnName].columnId,
      ),
      1: groupHeadings,
      3: headers,
      4: dataRow,
      ...options.extraRows,
    }),
    table: {
      name: configSheetFloorSeed.spreadsheetConfig.tableName,
      startRowIndex: tableHeaderRowIndex,
      startColumnIndex: startTableColIndex,
      endRowIndex: options.tableEndRowIndex ?? 5,
      endColumnIndex: sscColumns.length,
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
      endRowIndex: table === "header-only" ? headerOnlyTableEndRowIndex : 5,
    },
  };
}

function seedFixture(
  options: {
    idDelimiter?: string;
    idHeader?: string;
    testColumnId?: string;
    fillRowIdsRunStatusColumnId?: string;
    spreadsheetConfigTableEndRowIndex?: number;
    spreadsheetConfigExtraRows?: Record<
      number,
      readonly (string | number | boolean | null)[]
    >;
    valueConfigSheet?: FakeSheetProperties;
    extraSheets?: FakeSheetProperties[];
    extraSheetConfigDataRows?: Record<
      number,
      readonly (string | number | boolean | null)[]
    >;
    sheetConfigTableEndRowIndex?: number;
    startTableColumnIndexBase1?: number;
    spreadsheetConfigProtections?: GoogleAppsScript.Sheets.Schema.ProtectedRange[];
  } = {},
) {
  const idDelimiter = options.idDelimiter ?? ":";
  const testColumnId = options.testColumnId ?? "c:itm:xyz123";
  return stubSheetsService({
    sheets: [
      spreadsheetConfigSheet(idDelimiter, {
        idHeader: options.idHeader,
        tableEndRowIndex: options.spreadsheetConfigTableEndRowIndex,
        fillRowIdsRunStatusColumnId: options.fillRowIdsRunStatusColumnId,
        extraRows: options.spreadsheetConfigExtraRows,
        startTableColumnIndexBase1: options.startTableColumnIndexBase1,
        protectedRanges: options.spreadsheetConfigProtections,
      }),
      {
        sheetId: sheetConfigGid,
        title: "Sheet Config",
        rows: buildGridRows({
          0: [
            sc.sheetGid.columnId,
            sc.sheetTitle.columnId,
            sc.letApiAccess.columnId,
          ],
          3: [sc.sheetGid.header, sc.sheetTitle.header, sc.letApiAccess.header],
          4: testSheetConfigRowWithApiAccess,
          ...options.extraSheetConfigDataRows,
        }),
        table: {
          name: configSheetFloorSeed.sheetConfig.tableName,
          startRowIndex: tableHeaderRowIndex,
          startColumnIndex: startTableColIndex,
          endRowIndex: options.sheetConfigTableEndRowIndex ?? 5,
          endColumnIndex: startTableColIndex + 3,
          columnTypes: columnTypesByHeader("sheetConfig", [
            sc.sheetGid.header,
            sc.sheetTitle.header,
            sc.letApiAccess.header,
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
            cc.sheetGid.columnId,
            cc.columnId.columnId,
            cc.sheetTitle.columnId,
            cc.header.columnId,
            cc.emptyValueAllowed.columnId,
          ],
          3: [
            cc.sheetGid.header,
            cc.columnId.header,
            cc.sheetTitle.header,
            cc.header.header,
            cc.emptyValueAllowed.header,
          ],
        }),
        table: {
          name: configSheetFloorSeed.columnConfig.tableName,
          startRowIndex: tableHeaderRowIndex,
          startColumnIndex: startTableColIndex,
          endRowIndex: 5,
          endColumnIndex: startTableColIndex + 5,
          columnTypes: {
            0: "DOUBLE",
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
          0: [testColumnId],
          3: ["Some Header"],
          4: [],
        }),
        table: { endRowIndex: 5 },
      },
      options.valueConfigSheet ?? floorValueConfigTab(),
      ...(options.extraSheets ?? []),
    ],
  });
}

const valueConfigFloorGid = getSheetTraitByName("valueConfig", "sheetGid");

function floorValueConfigTab(): FakeSheetProperties {
  return {
    ...valueConfigTab({
      sheetId: valueConfigFloorGid,
      title: "Value Config",
      columnId: "c:vcf:abc1234",
    }),
    table: {
      name: configSheetFloorSeed.valueConfig.tableName,
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
    table: { endRowIndex: 5 },
  };
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

function floorWarningDescriptions(
  requests: GoogleAppsScript.Sheets.Schema.Request[] | undefined,
): string[] {
  return (requests ?? []).flatMap((request) => {
    const description = request.addProtectedRange?.protectedRange?.description;
    return description === undefined ? [] : [description];
  });
}

function driftedFloorWarningDescription(): string {
  return "Config-sheet floor · Spreadsheet Config · warning";
}

function sheetConfigLetApiAccess(
  orchestrator: ConfigCoordinator,
  sheetGid: number,
): boolean | "" {
  const sheet = orchestrator.sheetConfigOperator.sheet;
  const col = sheet.columns("sheetGid", "letApiAccess");
  const rowIndex = sheet.rowIndexesActiveWithData.find(
    (index) => col.sheetGid.value(index) === sheetGid,
  );
  if (rowIndex === undefined) {
    throw new Error(`No Sheet Config row for gid ${sheetGid}`);
  }
  return col.letApiAccess.valueOrEmpty(rowIndex);
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
  it("flushes the config-sheet floor, then Sheet Config and Column Config changes", () => {
    const { batchUpdateCalls } = seedFixture();

    const orchestrator = ConfigCoordinator.init();
    orchestrator.syncAndFlushConfigSheets();

    expect(batchUpdateCalls.length).toBe(2);
    expect(floorWarningDescriptions(batchUpdateCalls[0]?.requests)).toContain(
      driftedFloorWarningDescription(),
    );
    expect(batchUpdateCalls[1]?.requests?.length).toBeGreaterThan(0);
    // The gathered column ID rode the appended Column Config row into the flush.
    expect(orchestrator.sheetConfigOperator.newSheetConfigs().item).toEqual({
      sheetGid: testSheetGid,
      idPrefix: "itm",
      hasIdColumn: false,
      hasNameColumn: false,
    });
  });

  it("returns the untyped-column summary for the endpoint to report, with no file sources", () => {
    seedFixture();

    const summary = ConfigCoordinator.init().syncConfigSheetRows();

    expect(summary).toContain("1 column(s) across 1 sheet(s)");
    expect(typeof summary).toBe("string");
  });

  it("flushes floor warnings before the live Spreadsheet Config overlay is read", () => {
    const { batchUpdateCalls } = seedFixture({
      startTableColumnIndexBase1: 2,
    });

    expect(() => ConfigCoordinator.init().syncAndFlushConfigSheets()).toThrow();

    expect(floorWarningDescriptions(batchUpdateCalls[0]?.requests)).toContain(
      driftedFloorWarningDescription(),
    );
  });

  it("returns the floor report beside the untyped-column summary, as one line", () => {
    seedFixture({
      spreadsheetConfigProtections: [driftedFloorWarningProtection()],
    });

    const summary = ConfigCoordinator.init().syncConfigSheetRows();

    expect(summary).toContain("Replaced drifted:");
    expect(summary).toContain(driftedFloorWarningDescription());
    expect(summary).toContain("1 column(s) across 1 sheet(s)");
    expect(summary).not.toContain("\n");
  });
});

describe("ConfigCoordinator.generateConfigFiles", () => {
  it("returns every file's source together, reflecting the synced state", () => {
    seedFixture();

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    expect(typeof parsed.spreadsheetConfig).toBe("string");
    expect(typeof parsed.sheetConfigs).toBe("string");
    expect(typeof parsed.columnConfigs).toBe("string");
    expect(parsed.sheetConfigs).toContain('"item"');
    expect(typeof parsed.valueConfigs).toBe("string");
    // Gathered into Column Config by the sync, then headed by _updateProgrammaticValues.
    expect(parsed.columnConfigs).toContain("c:itm:xyz123");
  });

  it("refuses a live ID delimiter change, naming it with the expected one", () => {
    seedFixture({ idDelimiter: "|", testColumnId: "" });

    expect(() =>
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs"),
    ).toThrow('Spreadsheet Config column "ID delimiter" is "|"; expected ":".');
  });

  it("emits the live Spreadsheet Config values", () => {
    seedFixture({ idHeader: "Key" });

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    expect(parsed.spreadsheetConfig).toContain('idHeader: "Key"');
  });

  it("clears the live layout after the call returns", () => {
    seedFixture({ idHeader: "Key" });

    ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    expect(ssConfigGet("idHeader")).toBe(spreadsheetConfig.idHeader);
  });

  it("clears the live layout when later work throws", () => {
    stubSheetsService({
      sheets: [spreadsheetConfigSheet(":", { idHeader: "Key" })],
    });

    expect(() =>
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs"),
    ).toThrow(/need a full row\/column fetch but have no Table object/);
    expect(ssConfigGet("idHeader")).toBe(spreadsheetConfig.idHeader);
  });

  it("carries the untyped-column summary back, since no run status cell will show it", () => {
    seedFixture();

    expect(
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs")
        .untypedColumnsSummary,
    ).toContain("1 column(s) across 1 sheet(s)");
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
      "1 column(s) across 1 sheet(s)",
    );
  });

  it("carries the declared-cell report back, since no run status cell will show it", () => {
    seedFixture({
      extraSheetConfigDataRows: {
        5: [spreadsheetConfigGid, "Spreadsheet Config", false, ""],
      },
      sheetConfigTableEndRowIndex: 6,
    });

    expect(
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs")
        .declaredCellReport,
    ).toContain("Sheet Config · Let api access · Spreadsheet Config → TRUE");
  });

  it("still catalogs value titles after Column Config pruned a stale row of its own", () => {
    const columnIdRow = [
      cc.sheetGid.columnId,
      cc.columnId.columnId,
      cc.sheetTitle.columnId,
      cc.header.columnId,
      cc.emptyValueAllowed.columnId,
    ];
    stubSheetsService({
      sheets: [
        spreadsheetConfigSheet(":"),
        {
          sheetId: sheetConfigGid,
          title: "Sheet Config",
          rows: buildGridRows({
            0: [
              sc.sheetGid.columnId,
              sc.sheetTitle.columnId,
              sc.letApiAccess.columnId,
            ],
            3: [
              sc.sheetGid.header,
              sc.sheetTitle.header,
              sc.letApiAccess.header,
            ],
            4: [columnConfigGid, "Column Config", true],
          }),
          table: {
            name: configSheetFloorSeed.sheetConfig.tableName,
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
              cc.sheetGid.header,
              cc.columnId.header,
              cc.sheetTitle.header,
              cc.header.header,
              cc.emptyValueAllowed.header,
            ],
            4: [
              columnConfigGid,
              cc.sheetGid.columnId,
              "Column Config",
              cc.sheetGid.header,
            ],
            5: [columnConfigGid, "c:ccf:stale-gone", "Column Config", "Gone"],
          }),
          table: {
            name: configSheetFloorSeed.columnConfig.tableName,
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

  describe("floor identity", () => {
    const movedValueConfigGid = 999000111;

    function seedValueConfigTab(tab: FakeSheetProperties) {
      seedFixture({
        valueConfigSheet: tab,
        extraSheetConfigDataRows: { 5: [tab.sheetId, tab.title, true, ""] },
        sheetConfigTableEndRowIndex: 6,
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

    it("fails a floor column's changed column ID with the identity guard's message, before the floor seed check", () => {
      seedFixture({ fillRowIdsRunStatusColumnId: "c:sscf:moved01" });

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
        'Sheet "Item" sampled ID prefix "zzz" differs from last generated "itm".',
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
      extraSheetConfigDataRows: {
        5: [draftGid, draftTitle, false, ""],
      },
      sheetConfigTableEndRowIndex: 6,
    });

    const orchestrator = ConfigCoordinator.init();
    expect(() => orchestrator.syncConfigSheetRows()).not.toThrow();
    expect(
      orchestrator.sheetConfigOperator.newSheetConfigs().addWidgetOrder,
    ).toBeUndefined();
    expect(
      orchestrator.sheetConfigOperator.sheet
        .column("sheetGid")
        .hasValue(draftGid),
    ).toBe(true);
  });

  it("aborts and names the tab when Let api access is on and the Table has no data row", () => {
    seedFixture({
      extraSheets: [headerOnlyDraftSheet()],
      extraSheetConfigDataRows: {
        5: [draftGid, draftTitle, true, ""],
      },
      sheetConfigTableEndRowIndex: 6,
    });

    expect(() => ConfigCoordinator.init().syncConfigSheetRows()).toThrow(
      /Add Widget Order.*at least one data row/,
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
      extraSheetConfigDataRows: {
        5: [draftGid, draftTitle, true, "awo"],
      },
      sheetConfigTableEndRowIndex: 6,
    });

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    expect(parsed.sheetConfigs).toContain(
      `"addWidgetOrder": { "sheetGid": ${draftGid}, "idPrefix": "awo", "hasIdColumn": true, "hasNameColumn": true }`,
    );
  });

  it("treats a never-ticked Let api access box as off", () => {
    seedFixture({
      extraSheets: [headerOnlyDraftSheet()],
      extraSheetConfigDataRows: {
        5: [draftGid, draftTitle, null, ""],
      },
      sheetConfigTableEndRowIndex: 6,
    });

    const orchestrator = ConfigCoordinator.init();
    expect(() => orchestrator.syncConfigSheetRows()).not.toThrow();
    expect(
      orchestrator.sheetConfigOperator.newSheetConfigs().addWidgetOrder,
    ).toBeUndefined();
  });

  it("catalogues a brand-new header-only tab with Let api access off", () => {
    seedFixture({ extraSheets: [headerOnlyDraftSheet()] });

    const orchestrator = ConfigCoordinator.init();
    const parsed = orchestrator.generateConfigFiles("../makeConfigs");
    expect(
      orchestrator.sheetConfigOperator.sheet
        .column("sheetGid")
        .hasValue(draftGid),
    ).toBe(true);
    expect(parsed.sheetConfigs).not.toContain("addWidgetOrder");
  });

  it("omits a draft that has data rows until Let api access is ticked", () => {
    seedFixture({
      extraSheets: [headerOnlyDraftSheet({ table: "with-data-row" })],
      extraSheetConfigDataRows: {
        5: [draftGid, draftTitle, false, "awo"],
      },
      sheetConfigTableEndRowIndex: 6,
    });

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    expect(parsed.sheetConfigs).toContain('"item"');
    expect(parsed.sheetConfigs).not.toContain("addWidgetOrder");
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
    const gids = orchestrator.sheetConfigOperator.sheet.column("sheetGid");
    expect(gids.hasValue(draftGid)).toBe(true);
    expect(gids.hasValue(secondDraftGid)).toBe(true);
  });

  it("regens a healthy Let api access sheet while cataloguing an empty draft", () => {
    seedFixture({ extraSheets: [headerOnlyDraftSheet()] });

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    expect(parsed.sheetConfigs).toContain('"item"');
    expect(parsed.sheetConfigs).not.toContain("addWidgetOrder");
    expect(parsed.columnConfigs).toContain("c:itm:xyz123");
  });

  it("still syncs the config-describing sheets when they have Let api access", () => {
    seedFixture({
      extraSheetConfigDataRows: {
        5: [sheetConfigGid, "Sheet Config", true, "scf"],
        6: [columnConfigGid, "Column Config", true, "ccf"],
      },
      sheetConfigTableEndRowIndex: 7,
    });

    expect(() => ConfigCoordinator.init().syncConfigSheetRows()).not.toThrow();
  });

  it("writes an unticked Let api access on a floor tab back to TRUE and names the tab", () => {
    seedFixture({
      extraSheetConfigDataRows: {
        5: [spreadsheetConfigGid, "Spreadsheet Config", false, ""],
      },
      sheetConfigTableEndRowIndex: 6,
    });

    const orchestrator = ConfigCoordinator.init();
    const report = orchestrator.syncConfigSheetRows();

    expect(sheetConfigLetApiAccess(orchestrator, spreadsheetConfigGid)).toBe(
      true,
    );
    expect(report).toContain(
      "Sheet Config · Let api access · Spreadsheet Config → TRUE",
    );
  });

  it("appends a missing floor-tab row with Let api access TRUE", () => {
    seedFixture();

    const orchestrator = ConfigCoordinator.init();
    const report = orchestrator.syncConfigSheetRows();

    expect(
      orchestrator.sheetConfigOperator.sheet
        .column("sheetGid")
        .hasValue(spreadsheetConfigGid),
    ).toBe(true);
    expect(sheetConfigLetApiAccess(orchestrator, spreadsheetConfigGid)).toBe(
      true,
    );
    expect(report).toContain(
      "Sheet Config · Let api access · Spreadsheet Config → TRUE",
    );
  });

  it("leaves a non-floor tab's unticked Let api access alone", () => {
    seedFixture({
      extraSheets: [headerOnlyDraftSheet({ table: "with-data-row" })],
      extraSheetConfigDataRows: {
        5: [draftGid, draftTitle, false, "awo"],
      },
      sheetConfigTableEndRowIndex: 6,
    });

    const orchestrator = ConfigCoordinator.init();
    const report = orchestrator.syncConfigSheetRows();

    expect(sheetConfigLetApiAccess(orchestrator, draftGid)).toBe(false);
    expect(report ?? "").not.toContain(draftTitle);
  });

  it("fails when a Let api access sheet has no Table", () => {
    seedFixture({
      extraSheets: [headerOnlyDraftSheet({ table: "missing" })],
      extraSheetConfigDataRows: {
        5: [draftGid, draftTitle, true, "awo"],
      },
      sheetConfigTableEndRowIndex: 6,
    });

    expect(() => ConfigCoordinator.init().syncConfigSheetRows()).toThrow(
      /Active table is null/,
    );
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
      table: { endRowIndex: 5 },
    };
  }

  function sheetConfigRow(props: {
    sheetId: number;
    title: string;
    leftoverIdPrefix: string;
  }) {
    return [props.sheetId, props.title, true, props.leftoverIdPrefix];
  }

  it("assigns a generated prefix to a new Let api access sheet and mints column IDs with it", () => {
    seedFixture({
      extraSheets: [letApiAccessSheet({ sheetId: widgetGid, title: "Widget" })],
      extraSheetConfigDataRows: {
        5: sheetConfigRow({
          sheetId: widgetGid,
          title: "Widget",
          leftoverIdPrefix: "zzzz",
        }),
      },
      sheetConfigTableEndRowIndex: 6,
    });

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    expect(parsed.sheetConfigs).toContain(
      `"widget": { "sheetGid": ${widgetGid}, "idPrefix": "wdg", "hasIdColumn": false, "hasNameColumn": true }`,
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
      extraSheetConfigDataRows: {
        5: sheetConfigRow({
          sheetId: widgetGid,
          title: "Renamed Tab",
          leftoverIdPrefix: "zzzz",
        }),
      },
      sheetConfigTableEndRowIndex: 6,
    });

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    expect(parsed.sheetConfigs).toContain(
      `"renamedTab": { "sheetGid": ${widgetGid}, "idPrefix": "wdg", "hasIdColumn": false, "hasNameColumn": true }`,
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
      extraSheetConfigDataRows: {
        5: sheetConfigRow({
          sheetId: widgetGid,
          title: "Widget",
          leftoverIdPrefix: "zzzz",
        }),
      },
      sheetConfigTableEndRowIndex: 6,
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
      extraSheetConfigDataRows: {
        5: sheetConfigRow({
          sheetId: widgetGid,
          title: "Widget",
          leftoverIdPrefix: "wdg",
        }),
      },
      sheetConfigTableEndRowIndex: 6,
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
      extraSheetConfigDataRows: {
        5: sheetConfigRow({
          sheetId: widgetGid,
          title: "Widget",
          leftoverIdPrefix: "wdg",
        }),
        6: sheetConfigRow({
          sheetId: otherGid,
          title: "Gizmo",
          leftoverIdPrefix: "gzm",
        }),
      },
      sheetConfigTableEndRowIndex: 7,
    });

    expect(() =>
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs"),
    ).toThrow(/Widget.*Gizmo.*"wdg"/);
  });

  it("reports a sampled prefix that differs from the last generated sheet configs without failing", () => {
    seedFixture({ testColumnId: "c:zzz:xyz123" });

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    expect(parsed.idPrefixReport).toBe(
      'Sheet "Item" sampled ID prefix "zzz" differs from last generated "itm".',
    );
    expect(parsed.sheetConfigs).toContain('"idPrefix": "zzz"');
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
      extraSheetConfigDataRows: {
        5: [widgetGid, "Widget", false, "wdg"],
      },
      sheetConfigTableEndRowIndex: 6,
    });

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    expect(parsed.sheetConfigs).not.toContain("widget");
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
      extraSheetConfigDataRows: {
        5: sheetConfigRow({
          sheetId: widgetGid,
          title: "Gadget",
          leftoverIdPrefix: "gdg",
        }),
        6: sheetConfigRow({
          sheetId: otherGid,
          title: "Widget",
          leftoverIdPrefix: "zzzz",
        }),
      },
      sheetConfigTableEndRowIndex: 7,
    });

    const parsed =
      ConfigCoordinator.init().generateConfigFiles("../makeConfigs");
    expect(parsed.sheetConfigs).toContain(
      `"gadget": { "sheetGid": ${widgetGid}, "idPrefix": "wdg", "hasIdColumn": false, "hasNameColumn": true }`,
    );
    expect(parsed.sheetConfigs).toContain(
      `"widget": { "sheetGid": ${otherGid}, "idPrefix": "wdgt", "hasIdColumn": false, "hasNameColumn": true }`,
    );
  });
});
