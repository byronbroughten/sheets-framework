import { beforeEach, describe, expect, it } from "vitest";

import type { ModelableEditProtection } from "../00_Source/RawSource/EditProtection";
import { installedConfigs } from "../01_SpreadsheetSchema/configRegister";
import {
  configSheetFloorSeed,
  floorSeedColumns,
} from "../01_SpreadsheetSchema/configSheetFloorSeed";
import { getSheetTraitByName } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import { sheetLayout } from "../01_SpreadsheetSchema/sheetLayout";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  type FakeCell,
  type FakeSheetProperties,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { ConfigSheetFloor } from "./ConfigSheetFloor";
import { floorSheetNames } from "./ConfigSheetFloor/floorSeedLookups";
import { floorRecreatableColumns } from "./ConfigSheetFloor/FloorTabColumnCreator";
import { selfDescribingRowColumns } from "./ConfigSheetFloor/FloorTabEditWarning";

const { columnConfigs } = installedConfigs();

const spreadsheetConfigGid = getSheetTraitByName(
  "spreadsheetConfig",
  "sheetGid",
);
const sheetConfigGid = getSheetTraitByName("sheetConfig", "sheetGid");
const columnConfigGid = getSheetTraitByName("columnConfig", "sheetGid");
const valueConfigGid = getSheetTraitByName("valueConfig", "sheetGid");
const actionRowIndex = sheetLayout.actionRowIndex;
const topDataRowIndex = sheetLayout.tableHeaderRowIndex + 1;
const floorWarningPrefix = "Config-sheet floor";
const ssc = columnConfigs.spreadsheetConfig;
const sc = columnConfigs.sheetConfig;
const cc = columnConfigs.columnConfig;

const sscColumns = [
  "tableMenuSpace",
  "fillRowIdsTimeLastRan",
  "fillRowIdsRunStatus",
  "syncConfigSheetRowsTimeLastRan",
  "syncConfigSheetRowsRunStatus",
] as const;

beforeEach(() => {
  stubLogger();
});

function sscField<K extends (typeof sscColumns)[number]>(columnName: K) {
  return ssc[columnName];
}

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
  overrides: Record<string, string> = {},
): Record<number, string> {
  const types: Record<number, string> = {};
  headers.forEach((header, colIndex) => {
    const columnType = overrides[header] ?? floorSeedType(sheetName, header);
    if (columnType !== undefined) types[colIndex] = columnType;
  });
  return types;
}

function spreadsheetConfigFixtureColumnTypes(
  sscOrder: readonly (typeof sscColumns)[number][],
  sscHeaders: readonly string[],
  options: {
    columnTypesAreUnset?: boolean;
    spreadsheetConfigColumnTypes?: Partial<
      Record<(typeof sscColumns)[number], string>
    >;
  },
): Record<number, string> {
  if (options.columnTypesAreUnset) {
    const types: Record<number, string> = {};
    sscOrder.forEach((columnName, colIndex) => {
      const columnType = options.spreadsheetConfigColumnTypes?.[columnName];
      if (columnType !== undefined) types[colIndex] = columnType;
    });
    return types;
  }
  const typeOverrides = Object.fromEntries(
    Object.entries(options.spreadsheetConfigColumnTypes ?? {}).map(
      ([columnName, columnType]) => [
        sscField(columnName as (typeof sscColumns)[number]).header,
        columnType,
      ],
    ),
  );
  return columnTypesByHeader("spreadsheetConfig", sscHeaders, typeOverrides);
}

function matchingFloorColumnTypes(
  sheetName: "sheetConfig" | "columnConfig",
  headers: readonly string[],
  columnTypesAreUnset: boolean | undefined,
): Record<number, string> | undefined {
  if (columnTypesAreUnset) return undefined;
  return columnTypesByHeader(sheetName, headers);
}

function sheetAbsoluteTypes(
  types: Record<number, string> | undefined,
  startTableColIndex: number,
): Record<number, string> | undefined {
  if (types === undefined || startTableColIndex === 0) return types;
  return Object.fromEntries(
    Object.entries(types).map(([colIndex, columnType]) => [
      Number(colIndex) + startTableColIndex,
      columnType,
    ]),
  );
}

function padLeadingColumns(
  row: readonly FakeCell[],
  startTableColIndex: number,
): FakeCell[] {
  return [...Array.from({ length: startTableColIndex }, () => null), ...row];
}

function withExtraColumn<T>(row: T[], extra: T | undefined): T[] {
  if (extra === undefined) return row;
  return [...row, extra];
}

type FloorGrid = ReturnType<typeof stubSheetsService>["grid"];

// A sheet's single Table's columns in order, as header and column type.
function tableColumns(grid: FloorGrid, sheetGid: number) {
  return (grid.sheet(sheetGid).tables[0]?.columnProperties ?? []).map(
    ({ columnName, columnType }) => ({ columnName, columnType }),
  );
}

function seedTableColumns(
  columns: readonly { header: string; columnType: string }[],
) {
  return columns.map(({ header, columnType }) => ({
    columnName: header,
    columnType,
  }));
}

function floorTypedColumns(
  sheetName: "spreadsheetConfig" | "sheetConfig" | "columnConfig",
  headers: readonly string[],
) {
  return headers.map((header) => ({
    columnName: header,
    columnType: floorSeedType(sheetName, header),
  }));
}

function rowValues(
  grid: FloorGrid,
  sheetGid: number,
  rowIndex: number,
  endColumnIndex?: number,
) {
  return grid.sheet(sheetGid).values({
    startRowIndex: rowIndex,
    endRowIndex: rowIndex + 1,
    endColumnIndex,
  })[0];
}

function sheetSnapshot(grid: FloorGrid, sheetGid: number) {
  const sheet = grid.sheet(sheetGid);
  return {
    title: sheet.title,
    rows: sheet.rows(),
    tables: sheet.tables,
    protectedRanges: sheet.protectedRanges,
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

const sheetConfigColumns = ["sheetGid", "sheetTitle", "letApiAccess"] as const;
const columnConfigColumns = [
  "sheetGid",
  "columnId",
  "sheetTitle",
  "header",
  "emptyValueAllowed",
] as const;

const defaultSpreadsheetConfigHeaders = sscColumns.map(
  (columnName) => ssc[columnName].header,
);
const defaultSheetConfigHeaders = sheetConfigColumns.map(
  (columnName) => sc[columnName].header,
);
const defaultColumnConfigHeaders = columnConfigColumns.map(
  (columnName) => cc[columnName].header,
);

const restoredSpreadsheetConfigDataRow = sscColumns.map((columnName) =>
  columnName === "tableMenuSpace" ? "Not used" : "",
);

const businessSheetGid = 9001;
const addedColumnConfigColumn = { columnId: "c:ccf:notes", header: "Notes" };
const defaultSheetConfigGids = [
  businessSheetGid,
  spreadsheetConfigGid,
  sheetConfigGid,
  columnConfigGid,
  valueConfigGid,
];
const defaultColumnConfigRows = [
  { sheetGid: businessSheetGid, columnId: "c:biz:one" },
  { sheetGid: sheetConfigGid, columnId: sc.letApiAccess.columnId },
  { sheetGid: businessSheetGid, columnId: "c:biz:two" },
  { sheetGid: columnConfigGid, columnId: cc.emptyValueAllowed.columnId },
  { sheetGid: columnConfigGid, columnId: cc.header.columnId },
];

function dataRowsFrom(
  topRowIndex: number,
  rows: readonly FakeCell[][],
): Record<number, FakeCell[]> {
  return Object.fromEntries(rows.map((row, i) => [topRowIndex + i, row]));
}

function floorFixture(
  options: {
    spreadsheetConfigColumnOrder?: readonly (typeof sscColumns)[number][];
    spreadsheetConfigHeaders?: Partial<
      Record<(typeof sscColumns)[number], string>
    >;
    spreadsheetConfigColumnIds?: Partial<
      Record<(typeof sscColumns)[number], string>
    >;
    spreadsheetConfigGroupHeadings?: Partial<
      Record<(typeof sscColumns)[number], string>
    >;
    spreadsheetConfigColumnTypes?: Partial<
      Record<(typeof sscColumns)[number], string>
    >;
    columnTypesAreUnset?: boolean;
    tableMenuSpaceValue?: string;
    spreadsheetConfigProtections?: GoogleAppsScript.Sheets.Schema.ProtectedRange[];
    sheetConfigProtections?: GoogleAppsScript.Sheets.Schema.ProtectedRange[];
    columnConfigProtections?: GoogleAppsScript.Sheets.Schema.ProtectedRange[];
    sheetConfigColumnOrder?: readonly (typeof sheetConfigColumns)[number][];
    columnConfigColumnOrder?: readonly (typeof columnConfigColumns)[number][];
    sheetConfigHeaders?: Partial<Record<keyof typeof sc, string>>;
    sheetConfigColumnIds?: Partial<Record<keyof typeof sc, string>>;
    columnConfigColumnIds?: Partial<Record<keyof typeof cc, string>>;
    sheetConfigGids?: readonly number[];
    columnConfigRows?: readonly { sheetGid: number; columnId: string }[];
    spreadsheetConfigTitle?: string;
    spreadsheetConfigTableName?: string;
    valueConfig?: {
      title?: string;
      tableName?: string;
    };
    extraSheets?: FakeSheetProperties[];
    omitSheetGids?: readonly number[];
    isDryRun?: boolean;
    omitSpreadsheetConfigTable?: boolean;
    spreadsheetConfigExtraTables?: NonNullable<
      FakeSheetProperties["extraTables"]
    >;
    extraSpreadsheetConfigColumn?: {
      columnId: string;
      header: string;
      groupHeading?: string;
    };
    extraColumnConfigColumn?: { columnId: string; header: string };
  } = {},
) {
  const startTableColIndex = sheetLayout.startTableColIndex;
  const pad = (row: readonly FakeCell[]) =>
    padLeadingColumns(row, startTableColIndex);
  const sscOrder = options.spreadsheetConfigColumnOrder ?? sscColumns;
  const sscHeaders = sscOrder.map(
    (columnName) =>
      options.spreadsheetConfigHeaders?.[columnName] ??
      sscField(columnName).header,
  );
  const sscColumnIds = sscOrder.map(
    (columnName) =>
      options.spreadsheetConfigColumnIds?.[columnName] ??
      sscField(columnName).columnId,
  );
  const groupHeadings = sscOrder.map(
    (columnName) =>
      options.spreadsheetConfigGroupHeadings?.[columnName] ??
      spreadsheetConfigGroupHeading(sscField(columnName).header),
  );
  const extraColumn = options.extraSpreadsheetConfigColumn;
  const scOrder = options.sheetConfigColumnOrder ?? sheetConfigColumns;
  const ccOrder = options.columnConfigColumnOrder ?? columnConfigColumns;
  const ccHeaders = ccOrder.map((columnName) => cc[columnName].header);
  const extraCcColumn = options.extraColumnConfigColumn;
  const sheetConfigHeaders = scOrder.map(
    (columnName) =>
      options.sheetConfigHeaders?.[columnName] ?? sc[columnName].header,
  );
  const sheetConfigGids = options.sheetConfigGids ?? defaultSheetConfigGids;
  const columnConfigRows = options.columnConfigRows ?? defaultColumnConfigRows;
  const dataRow = sscOrder.map((columnName) =>
    columnName === "tableMenuSpace"
      ? (options.tableMenuSpaceValue ?? "Not used")
      : "",
  );

  return stubSheetsService({
    sheets: [
      {
        sheetId: spreadsheetConfigGid,
        title: options.spreadsheetConfigTitle ?? "Spreadsheet Config",
        rows: buildGridRows({
          0: pad(withExtraColumn(sscColumnIds, extraColumn?.columnId)),
          1: pad(
            withExtraColumn(groupHeadings, extraColumn?.groupHeading ?? ""),
          ),
          3: pad(withExtraColumn(sscHeaders, extraColumn?.header)),
          4: pad(withExtraColumn(dataRow, "")),
        }),
        table: options.omitSpreadsheetConfigTable
          ? undefined
          : {
              name:
                options.spreadsheetConfigTableName ??
                configSheetFloorSeed.spreadsheetConfig.tableName,
              startColumnIndex: startTableColIndex,
              endRowIndex: 5,
              endColumnIndex:
                startTableColIndex +
                sscOrder.length +
                (extraColumn === undefined ? 0 : 1),
              columnTypes: sheetAbsoluteTypes(
                spreadsheetConfigFixtureColumnTypes(
                  sscOrder,
                  sscHeaders,
                  options,
                ),
                startTableColIndex,
              ),
            },
        extraTables: options.spreadsheetConfigExtraTables,
        protectedRanges: options.spreadsheetConfigProtections,
      },
      {
        sheetId: sheetConfigGid,
        title: "Sheet Config",
        rows: buildGridRows({
          0: pad(
            scOrder.map(
              (columnName) =>
                options.sheetConfigColumnIds?.[columnName] ??
                sc[columnName].columnId,
            ),
          ),
          3: pad(sheetConfigHeaders),
          ...dataRowsFrom(
            topDataRowIndex,
            sheetConfigGids.map((gid) => {
              const cells = {
                sheetGid: gid,
                sheetTitle: `Tab ${gid}`,
                letApiAccess: true,
              };
              return pad(scOrder.map((columnName) => cells[columnName]));
            }),
          ),
        }),
        table: {
          name: configSheetFloorSeed.sheetConfig.tableName,
          startColumnIndex: startTableColIndex,
          endRowIndex: topDataRowIndex + sheetConfigGids.length,
          endColumnIndex: startTableColIndex + scOrder.length,
          columnTypes: sheetAbsoluteTypes(
            matchingFloorColumnTypes(
              "sheetConfig",
              sheetConfigHeaders,
              options.columnTypesAreUnset,
            ),
            startTableColIndex,
          ),
        },
        protectedRanges: options.sheetConfigProtections,
      },
      {
        sheetId: columnConfigGid,
        title: "Column Config",
        rows: buildGridRows({
          0: pad(
            withExtraColumn<FakeCell>(
              ccOrder.map(
                (columnName) =>
                  options.columnConfigColumnIds?.[columnName] ??
                  cc[columnName].columnId,
              ),
              extraCcColumn?.columnId,
            ),
          ),
          3: pad(withExtraColumn(ccHeaders, extraCcColumn?.header)),
          ...dataRowsFrom(
            topDataRowIndex,
            columnConfigRows.map(({ sheetGid, columnId }) => {
              const cells = {
                sheetGid,
                columnId,
                sheetTitle: "",
                header: "",
                emptyValueAllowed: false,
              };
              return pad(
                withExtraColumn<FakeCell>(
                  ccOrder.map((columnName) => cells[columnName]),
                  extraCcColumn === undefined ? undefined : "",
                ),
              );
            }),
          ),
        }),
        table: {
          name: configSheetFloorSeed.columnConfig.tableName,
          startColumnIndex: startTableColIndex,
          endRowIndex: topDataRowIndex + columnConfigRows.length,
          endColumnIndex:
            startTableColIndex +
            ccOrder.length +
            (extraCcColumn === undefined ? 0 : 1),
          columnTypes: sheetAbsoluteTypes(
            matchingFloorColumnTypes(
              "columnConfig",
              ccHeaders,
              options.columnTypesAreUnset,
            ),
            startTableColIndex,
          ),
        },
        protectedRanges: options.columnConfigProtections,
      },
      {
        sheetId: valueConfigGid,
        title:
          options.valueConfig?.title ?? configSheetFloorSeed.valueConfig.title,
        table: {
          name:
            options.valueConfig?.tableName ??
            configSheetFloorSeed.valueConfig.tableName,
          endRowIndex: 5,
        },
      },
      ...(options.extraSheets ?? []),
    ].filter((sheet) => !options.omitSheetGids?.includes(sheet.sheetId)),
    isDryRun: options.isDryRun,
  });
}

function applyFloor() {
  const floor = ConfigSheetFloor.init();
  const report = floor.ensure();
  floor.ss.batchUpdateGSheets();
  return { floor, report };
}

function protectionsOf(
  floor: ConfigSheetFloor,
  sheetName: "spreadsheetConfig" | "sheetConfig" | "columnConfig",
): ModelableEditProtection[] {
  const sheet = floor.ss.sheet(sheetName);
  sheet.prepFetchEditProtections();
  floor.ss.fetchAllPrepped({ skipFetchingProperties: true });
  return sheet
    .editProtections()
    .flatMap((protection) =>
      protection.kind === "unmodelable" ? [] : [protection],
    );
}

function floorWarningDescription(title: string): string {
  return `${floorWarningPrefix} · ${title} · warning`;
}

const spreadsheetConfigEditableRanges = [
  {
    sheetId: spreadsheetConfigGid,
    startRowIndex: actionRowIndex,
    endRowIndex: actionRowIndex + 1,
    startColumnIndex: 1,
    endColumnIndex: 2,
  },
  {
    sheetId: spreadsheetConfigGid,
    startRowIndex: actionRowIndex,
    endRowIndex: actionRowIndex + 1,
    startColumnIndex: 3,
    endColumnIndex: 4,
  },
];
const sheetConfigEditableRanges = [
  {
    sheetId: sheetConfigGid,
    startRowIndex: topDataRowIndex,
    endRowIndex: topDataRowIndex + 1,
    startColumnIndex: 2,
    endColumnIndex: 3,
  },
  {
    sheetId: sheetConfigGid,
    startRowIndex: topDataRowIndex + defaultSheetConfigGids.length,
    startColumnIndex: 2,
    endColumnIndex: 3,
  },
];
const columnConfigEditableRanges = [
  {
    sheetId: columnConfigGid,
    startRowIndex: topDataRowIndex,
    endRowIndex: topDataRowIndex + 1,
    startColumnIndex: 4,
    endColumnIndex: 5,
  },
  {
    sheetId: columnConfigGid,
    startRowIndex: topDataRowIndex + 2,
    endRowIndex: topDataRowIndex + 3,
    startColumnIndex: 4,
    endColumnIndex: 5,
  },
  {
    sheetId: columnConfigGid,
    startRowIndex: topDataRowIndex + 5,
    startColumnIndex: 4,
    endColumnIndex: 5,
  },
];

describe("ConfigSheetFloor", () => {
  it("puts one whole-sheet warning on each of Spreadsheet Config, Sheet Config and Column Config, with editable ranges where an edit sticks, and locks none", () => {
    const { grid } = floorFixture();
    const { floor } = applyFloor();

    const spreadsheet = protectionsOf(floor, "spreadsheetConfig");
    const sheetConfig = protectionsOf(floor, "sheetConfig");
    const columnConfig = protectionsOf(floor, "columnConfig");
    const all = [...spreadsheet, ...sheetConfig, ...columnConfig];

    expect(all).toHaveLength(3);
    expect(all.every((protection) => protection.kind === "warning")).toBe(true);
    expect(all.some((protection) => protection.kind === "lock")).toBe(false);

    expect(spreadsheet[0]).toMatchObject({
      description: floorWarningDescription("Spreadsheet Config"),
      range: { sheetId: spreadsheetConfigGid },
    });
    expect(spreadsheet[0]?.unprotectedRanges).toEqual(
      spreadsheetConfigEditableRanges,
    );
    expect(sheetConfig[0]).toMatchObject({
      description: floorWarningDescription("Sheet Config"),
      range: { sheetId: sheetConfigGid },
    });
    expect(sheetConfig[0]?.unprotectedRanges).toEqual(
      sheetConfigEditableRanges,
    );
    expect(columnConfig[0]).toMatchObject({
      description: floorWarningDescription("Column Config"),
      range: { sheetId: columnConfigGid },
    });
    expect(columnConfig[0]?.unprotectedRanges).toEqual(
      columnConfigEditableRanges,
    );
    expect(grid.sheet(valueConfigGid).protectedRanges).toEqual([]);
  });

  it("sends no batch update on a second run", () => {
    const { batchUpdateCount } = floorFixture();
    const { floor } = applyFloor();
    floor.ensure();
    floor.ss.batchUpdateGSheets();
    expect(batchUpdateCount()).toBe(1);
  });

  it("fetches the floor's grid in the same two reads it took before the carve read identity columns", () => {
    const { getByDataFilterCalls } = floorFixture();
    applyFloor();

    expect(getByDataFilterCalls).toHaveLength(2);
  });

  it("carves the four floor rows out of Sheet Config's Let api access column, leaving a bounded range above and an open-ended one below", () => {
    floorFixture();
    const { floor } = applyFloor();

    expect(protectionsOf(floor, "sheetConfig")[0]?.unprotectedRanges).toEqual(
      sheetConfigEditableRanges,
    );
  });

  it("carves Column Config's Empty value allowed at each floor column's row and leaves a business column's row editable", () => {
    floorFixture();
    const { floor } = applyFloor();

    expect(protectionsOf(floor, "columnConfig")[0]?.unprotectedRanges).toEqual(
      columnConfigEditableRanges,
    );
  });

  it("leaves Column Config's editable column open while no floor column has a row", () => {
    floorFixture({
      columnConfigRows: [{ sheetGid: businessSheetGid, columnId: "c:biz:one" }],
    });
    const { floor } = applyFloor();

    expect(protectionsOf(floor, "columnConfig")[0]?.unprotectedRanges).toEqual([
      {
        sheetId: columnConfigGid,
        startRowIndex: topDataRowIndex,
        startColumnIndex: 4,
        endColumnIndex: 5,
      },
    ]);
  });

  it("leaves a column added to Column Config editable beside the carved Empty value allowed", () => {
    floorFixture({ extraColumnConfigColumn: addedColumnConfigColumn });
    const { floor } = applyFloor();

    const [firstCarve, ...laterCarves] = columnConfigEditableRanges;
    expect(protectionsOf(floor, "columnConfig")[0]?.unprotectedRanges).toEqual([
      firstCarve,
      {
        sheetId: columnConfigGid,
        startRowIndex: topDataRowIndex,
        startColumnIndex: 5,
        endColumnIndex: 6,
      },
      ...laterCarves,
    ]);
  });

  it("keeps Column Config's editable column merged with an added one while no floor column has a row", () => {
    floorFixture({
      columnConfigRows: [{ sheetGid: businessSheetGid, columnId: "c:biz:one" }],
      extraColumnConfigColumn: addedColumnConfigColumn,
    });
    const { floor } = applyFloor();

    expect(protectionsOf(floor, "columnConfig")[0]?.unprotectedRanges).toEqual([
      {
        sheetId: columnConfigGid,
        startRowIndex: topDataRowIndex,
        startColumnIndex: 4,
        endColumnIndex: 6,
      },
    ]);
  });

  it("leaves a floor tab's row uncarved when Sheet Config holds none for it, and does not throw", () => {
    floorFixture({ sheetConfigGids: [businessSheetGid] });
    const { floor } = applyFloor();

    expect(protectionsOf(floor, "sheetConfig")[0]?.unprotectedRanges).toEqual([
      {
        sheetId: sheetConfigGid,
        startRowIndex: topDataRowIndex,
        startColumnIndex: 2,
        endColumnIndex: 3,
      },
    ]);
  });

  it("sends no batch update on a second sync against the grid the first one left", () => {
    const { batchUpdateCount } = floorFixture();
    applyFloor();
    expect(batchUpdateCount()).toBe(1);

    const second = applyFloor();

    expect(second.report).toBe("");
    expect(batchUpdateCount()).toBe(1);
  });

  it("covers a column added beside the floor and reports it", () => {
    floorFixture({
      extraSpreadsheetConfigColumn: {
        columnId: "c:sscf:notes",
        header: "Notes",
        groupHeading: "Mine",
      },
    });
    const { floor, report } = applyFloor();

    expect(report).toContain("Covered added columns:");
    expect(report).toContain("Spreadsheet Config · Notes");
    expect(
      protectionsOf(floor, "spreadsheetConfig")[0]?.unprotectedRanges,
    ).toEqual([
      ...spreadsheetConfigEditableRanges,
      {
        sheetId: spreadsheetConfigGid,
        startRowIndex: topDataRowIndex,
        startColumnIndex: sscColumns.length,
        endColumnIndex: sscColumns.length + 1,
      },
    ]);
  });

  it("replaces a drifted whole-sheet warning and reports it", () => {
    const driftedDescription = floorWarningDescription("Spreadsheet Config");
    floorFixture({
      spreadsheetConfigProtections: [
        {
          protectedRangeId: 41,
          description: driftedDescription,
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
        },
      ],
    });
    const { floor, report } = applyFloor();

    expect(report).toContain("Replaced drifted:");
    expect(report).toContain(driftedDescription);
    expect(
      protectionsOf(floor, "spreadsheetConfig")[0]?.unprotectedRanges,
    ).toEqual(spreadsheetConfigEditableRanges);
  });

  it("leaves a hand-set protection untouched", () => {
    const handSet = {
      protectedRangeId: 99,
      description: "Hand-set",
      warningOnly: true,
      range: {
        sheetId: sheetConfigGid,
        startRowIndex: 0,
        endRowIndex: 1,
        startColumnIndex: 0,
        endColumnIndex: 1,
      },
    };
    const { grid } = floorFixture({ sheetConfigProtections: [handSet] });
    const { floor } = applyFloor();

    expect(grid.sheet(sheetConfigGid).protectedRanges).toContainEqual(handSet);
    expect(protectionsOf(floor, "sheetConfig")).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ id: 99, description: "Hand-set" }),
        expect.objectContaining({
          description: floorWarningDescription("Sheet Config"),
        }),
      ]),
    );
  });

  it("removes leftover per-cell floor warnings", () => {
    const { grid } = floorFixture({
      spreadsheetConfigProtections: [
        {
          protectedRangeId: 41,
          description: `${floorWarningPrefix} · Spreadsheet Config · Table menu space (${ssc.tableMenuSpace.columnId}) · data · warning`,
          warningOnly: true,
          range: {
            sheetId: spreadsheetConfigGid,
            startRowIndex: topDataRowIndex,
            endRowIndex: topDataRowIndex + 1,
            startColumnIndex: 0,
            endColumnIndex: 1,
          },
        },
      ],
    });
    const { floor } = applyFloor();

    expect(
      grid
        .sheet(spreadsheetConfigGid)
        .protectedRanges.map((protection) => protection.protectedRangeId),
    ).not.toContain(41);
    expect(protectionsOf(floor, "spreadsheetConfig")).toHaveLength(1);
    expect(protectionsOf(floor, "spreadsheetConfig")[0]).toMatchObject({
      description: floorWarningDescription("Spreadsheet Config"),
      range: { sheetId: spreadsheetConfigGid },
    });
  });

  it("sets a floor column whose type differs from the seed and reports it", () => {
    const { grid } = floorFixture({
      spreadsheetConfigColumnTypes: { tableMenuSpace: "DOUBLE" },
    });
    const { report } = applyFloor();

    expect(report).toContain("Set column types:");
    expect(report).toContain(
      `Spreadsheet Config · Table menu space (${ssc.tableMenuSpace.columnId}) → TEXT`,
    );
    expect(tableColumns(grid, spreadsheetConfigGid)).toEqual(
      floorTypedColumns("spreadsheetConfig", defaultSpreadsheetConfigHeaders),
    );
    expect(tableColumns(grid, sheetConfigGid)).toEqual(
      floorTypedColumns("sheetConfig", defaultSheetConfigHeaders),
    );
    expect(tableColumns(grid, columnConfigGid)).toEqual(
      floorTypedColumns("columnConfig", defaultColumnConfigHeaders),
    );
  });

  it("leaves matching floor column types unchanged", () => {
    const { grid } = floorFixture();
    const { report } = applyFloor();

    expect(report).not.toContain("Set column types:");
    expect(tableColumns(grid, spreadsheetConfigGid)).toEqual(
      floorTypedColumns("spreadsheetConfig", defaultSpreadsheetConfigHeaders),
    );
    expect(tableColumns(grid, sheetConfigGid)).toEqual(
      floorTypedColumns("sheetConfig", defaultSheetConfigHeaders),
    );
    expect(tableColumns(grid, columnConfigGid)).toEqual(
      floorTypedColumns("columnConfig", defaultColumnConfigHeaders),
    );
  });

  it("does not set a type on a column added to Column Config", () => {
    const { grid } = floorFixture({
      columnTypesAreUnset: true,
      extraColumnConfigColumn: addedColumnConfigColumn,
    });
    applyFloor();

    expect(tableColumns(grid, columnConfigGid)).toEqual([
      ...floorTypedColumns("columnConfig", defaultColumnConfigHeaders),
      { columnName: addedColumnConfigColumn.header, columnType: undefined },
    ]);
  });

  it("sends no batch update on a second run after setting column types", () => {
    const { grid, batchUpdateCount } = floorFixture({
      columnTypesAreUnset: true,
    });
    applyFloor();

    expect(tableColumns(grid, sheetConfigGid)).toEqual(
      floorTypedColumns("sheetConfig", defaultSheetConfigHeaders),
    );
    applyFloor();
    expect(batchUpdateCount()).toBe(1);
  });

  it("renames a drifted floor tab back and reports it", () => {
    const { grid } = floorFixture({
      spreadsheetConfigTitle: "Old Spreadsheet Config",
    });
    const { report } = applyFloor();

    expect(report).toContain(
      'Restored tab titles: "Old Spreadsheet Config" → Spreadsheet Config',
    );
    expect(grid.sheet(spreadsheetConfigGid).title).toBe("Spreadsheet Config");
  });

  it("names a restored tab by its seed title in the rest of the report", () => {
    floorFixture({
      spreadsheetConfigTitle: "Old Spreadsheet Config",
      spreadsheetConfigHeaders: { tableMenuSpace: "Menu spacer" },
      spreadsheetConfigColumnTypes: { fillRowIdsRunStatus: "DOUBLE" },
      extraSpreadsheetConfigColumn: {
        columnId: "c:sscf:notes",
        header: "Notes",
      },
    });
    const { report } = applyFloor();

    expect(report).toContain(
      `Spreadsheet Config · Menu spacer (${ssc.tableMenuSpace.columnId}) → Table menu space`,
    );
    expect(report).toContain(
      `Spreadsheet Config · Fill row IDs, run status (${ssc.fillRowIdsRunStatus.columnId}) → TEXT`,
    );
    expect(report).toContain(
      "Covered added columns: Spreadsheet Config · Notes",
    );
    expect(report).not.toContain("Old Spreadsheet Config ·");
  });

  it("renames Value Config's tab back and reports it", () => {
    const { grid } = floorFixture({
      valueConfig: { title: "Values" },
    });
    const { report } = applyFloor();

    expect(report).toContain('Restored tab titles: "Values" → Value Config');
    expect(grid.sheet(valueConfigGid).title).toBe("Value Config");
  });

  it("returns the floor notice for a renamed Value Config off freshly fetched sheet properties and sends no batch update", () => {
    const { batchUpdateCount, getCalls } = floorFixture({
      valueConfig: { title: "Values" },
    });
    const notice = ConfigSheetFloor.init().changeNotice("other");

    expect(notice).toEqual({
      title: "Value Config is managed",
      message:
        'This tab keeps the name "Value Config". Your rename will switch back the next time configs sync.',
      untilClosed: false,
    });
    expect(getCalls).toHaveLength(1);
    expect(batchUpdateCount()).toBe(0);
  });

  it("renames a wrongly named floor Table and reports it", () => {
    const { grid } = floorFixture({
      spreadsheetConfigTableName: "wrongTable",
    });
    const { report } = applyFloor();

    expect(report).toContain(
      `Restored Table names: Spreadsheet Config's Table "wrongTable" → spreadsheetConfig`,
    );
    expect(grid.sheet(spreadsheetConfigGid).tables[0]?.name).toBe(
      "spreadsheetConfig",
    );
  });

  it("renames Value Config's Table back to valueConfig", () => {
    const { grid } = floorFixture({
      valueConfig: { tableName: "values" },
    });
    const { report } = applyFloor();

    expect(report).toContain(
      `Restored Table names: Value Config's Table "values" → valueConfig`,
    );
    expect(grid.sheet(valueConfigGid).tables[0]?.name).toBe("valueConfig");
  });

  it("throws naming the tab when a floor title sits on the wrong GID, and sends no batch update", () => {
    const { batchUpdateCount } = floorFixture({
      spreadsheetConfigTitle: "Old Spreadsheet Config",
      extraSheets: [
        {
          sheetId: 999001,
          title: "Spreadsheet Config",
          table: { endRowIndex: 5 },
        },
      ],
    });

    expect(() => applyFloor()).toThrow(
      'A tab titled "Spreadsheet Config" is not the floor tab.',
    );
    expect(batchUpdateCount()).toBe(0);
  });

  it("throws naming the tab when a floor tab has no Table, and sends no batch update", () => {
    const { batchUpdateCount } = floorFixture({
      omitSpreadsheetConfigTable: true,
    });

    expect(() => applyFloor()).toThrow(
      'Floor tab "Spreadsheet Config" has no Table.',
    );
    expect(batchUpdateCount()).toBe(0);
  });

  it("throws naming the tab when several Tables are present and none has the floor name, and sends no batch update", () => {
    const { batchUpdateCount } = floorFixture({
      spreadsheetConfigTableName: "firstWrong",
      spreadsheetConfigExtraTables: [{ endRowIndex: 5, name: "secondWrong" }],
    });

    expect(() => applyFloor()).toThrow(
      'Floor tab "Spreadsheet Config" has several Tables and none is named spreadsheetConfig.',
    );
    expect(batchUpdateCount()).toBe(0);
  });

  it("throws naming the tab when several Tables are present and one has the floor name, and sends no batch update", () => {
    const { batchUpdateCount } = floorFixture({
      spreadsheetConfigExtraTables: [{ endRowIndex: 5, name: "extra" }],
    });

    expect(() => applyFloor()).toThrow(
      '1 sheet(s) have more than one Table — delete the extras so each sheet has exactly one: "Spreadsheet Config"',
    );
    expect(batchUpdateCount()).toBe(0);
  });

  it("overwrites a drifted floor header and reports it", () => {
    const { grid } = floorFixture({
      spreadsheetConfigHeaders: { tableMenuSpace: "Menu spacer" },
    });
    const { report } = applyFloor();

    expect(report).toContain("Restored headers:");
    expect(report).toContain(
      `Spreadsheet Config · Menu spacer (${ssc.tableMenuSpace.columnId}) → Table menu space`,
    );
    expect(grid.sheet(spreadsheetConfigGid).cell(3, 0)).toBe(
      "Table menu space",
    );
  });

  it("overwrites a drifted floor column ID and reports it", () => {
    const { grid } = floorFixture({
      spreadsheetConfigColumnIds: { tableMenuSpace: "c:sscf:drifted" },
    });
    const { report } = applyFloor();

    expect(report).toContain("Restored column IDs:");
    expect(report).toContain(
      `Spreadsheet Config · Table menu space (c:sscf:drifted) → ${ssc.tableMenuSpace.columnId}`,
    );
    expect(grid.sheet(spreadsheetConfigGid).cell(0, 0)).toBe(
      ssc.tableMenuSpace.columnId,
    );
  });

  it("overwrites a drifted Sheet GID column ID on Sheet Config without throwing, and reports it", () => {
    const { grid } = floorFixture({
      sheetConfigColumnIds: { sheetGid: "c:scf:drifted" },
    });
    const { report } = applyFloor();

    expect(report).toContain(
      `Sheet Config · ${sc.sheetGid.header} (c:scf:drifted) → ${sc.sheetGid.columnId}`,
    );
    expect(grid.sheet(sheetConfigGid).cell(0, 0)).toBe(sc.sheetGid.columnId);
  });

  it("overwrites a drifted Column ID column ID on Column Config without throwing, and reports it", () => {
    const { grid } = floorFixture({
      columnConfigColumnIds: { columnId: "c:ccf:drifted" },
    });
    const { report } = applyFloor();

    expect(report).toContain(
      `Column Config · ${cc.columnId.header} (c:ccf:drifted) → ${cc.columnId.columnId}`,
    );
    expect(grid.sheet(columnConfigGid).cell(0, 1)).toBe(cc.columnId.columnId);
  });

  it("carves nothing on Sheet Config the sync its Sheet GID header has drifted, and restores the header without throwing", () => {
    const { grid } = floorFixture({
      sheetConfigHeaders: { sheetGid: "Tab gid" },
    });
    const { floor, report } = applyFloor();

    expect(protectionsOf(floor, "sheetConfig")[0]?.unprotectedRanges).toEqual([
      {
        sheetId: sheetConfigGid,
        startRowIndex: topDataRowIndex,
        startColumnIndex: 2,
        endColumnIndex: 3,
      },
    ]);
    expect(report).toContain(
      `Sheet Config · Tab gid (${sc.sheetGid.columnId}) → ${sc.sheetGid.header}`,
    );
    expect(grid.sheet(sheetConfigGid).cell(3, 0)).toBe(sc.sheetGid.header);
  });

  it("overwrites a drifted floor group heading and reports it", () => {
    const { grid } = floorFixture({
      spreadsheetConfigGroupHeadings: { fillRowIdsTimeLastRan: "Rules" },
    });
    const { report } = applyFloor();

    expect(report).toContain("Restored group headings:");
    expect(report).toContain(
      `Spreadsheet Config · Fill row IDs, time last ran (${ssc.fillRowIdsTimeLastRan.columnId}) → Fill Row IDs`,
    );
    expect(
      grid
        .sheet(spreadsheetConfigGid)
        .cell(1, sscColumns.indexOf("fillRowIdsTimeLastRan")),
    ).toBe("Fill Row IDs");
  });

  it("leaves an extra non-floor column's header, column ID and group heading alone", () => {
    const { grid } = floorFixture({
      extraSpreadsheetConfigColumn: {
        columnId: "c:sscf:notes",
        header: "Notes",
        groupHeading: "Mine",
      },
    });
    applyFloor();

    const extraColIndex = sscColumns.length;
    expect(
      grid.sheet(spreadsheetConfigGid).values({
        endRowIndex: topDataRowIndex + 1,
        startColumnIndex: extraColIndex,
        endColumnIndex: extraColIndex + 1,
      }),
    ).toEqual([["c:sscf:notes"], ["Mine"], [null], ["Notes"], [""]]);
  });

  it("writes Not used into a blank Table menu space data cell and reports it", () => {
    const { grid } = floorFixture({ tableMenuSpaceValue: "" });
    const { report } = applyFloor();

    expect(report).toContain('Restored Table menu space: "" → Not used');
    expect(
      rowValues(grid, spreadsheetConfigGid, topDataRowIndex, sscColumns.length),
    ).toEqual(restoredSpreadsheetConfigDataRow);
  });

  it("leaves a Table menu space data cell that already reads Not used as it is, and reports nothing about it", () => {
    const { grid } = floorFixture();
    const { report } = applyFloor();

    expect(report).not.toContain("Table menu space");
    expect(
      rowValues(grid, spreadsheetConfigGid, topDataRowIndex, sscColumns.length),
    ).toEqual(restoredSpreadsheetConfigDataRow);
  });

  it("overwrites an edited Table menu space data cell and reports it", () => {
    const { grid } = floorFixture({ tableMenuSpaceValue: "notes" });
    const { report } = applyFloor();

    expect(report).toContain('Restored Table menu space: "notes" → Not used');
    expect(grid.sheet(spreadsheetConfigGid).cell(topDataRowIndex, 0)).toBe(
      "Not used",
    );
  });

  it("recreates a missing Column Config Header column at the Table end with its header, column ID and heading, and reports it", () => {
    const { grid } = floorFixture({
      columnConfigColumnOrder: [
        "sheetGid",
        "columnId",
        "sheetTitle",
        "emptyValueAllowed",
      ],
    });
    const { report } = applyFloor();

    expect(
      rowValues(grid, columnConfigGid, sheetLayout.tableHeaderRowIndex),
    ).toEqual([
      cc.sheetGid.header,
      cc.columnId.header,
      cc.sheetTitle.header,
      cc.emptyValueAllowed.header,
      cc.header.header,
    ]);
    expect(
      grid.sheet(columnConfigGid).values({
        endRowIndex: sheetLayout.tableHeaderRowIndex,
        startColumnIndex: 4,
        endColumnIndex: 5,
      }),
    ).toEqual([[cc.header.columnId], [""], [null]]);
    expect(report).toContain(
      `Recreated columns: Column Config · ${cc.header.header} (${cc.header.columnId})`,
    );
  });

  it("recreates two missing columns on one tab at the Table end, in seed order", () => {
    const { grid } = floorFixture({
      columnConfigColumnOrder: ["sheetGid", "columnId", "emptyValueAllowed"],
    });
    applyFloor();

    expect(
      rowValues(grid, columnConfigGid, sheetLayout.tableHeaderRowIndex),
    ).toEqual([
      cc.sheetGid.header,
      cc.columnId.header,
      cc.emptyValueAllowed.header,
      cc.sheetTitle.header,
      cc.header.header,
    ]);
  });

  it("keeps each Column Config column type on its own column in a sync that recreates a column there", () => {
    const { grid } = floorFixture({
      columnTypesAreUnset: true,
      columnConfigColumnOrder: [
        "sheetGid",
        "columnId",
        "sheetTitle",
        "emptyValueAllowed",
      ],
    });
    applyFloor();

    const columns = tableColumns(grid, columnConfigGid);
    expect(columns.length).toBeGreaterThan(0);
    columns.forEach(({ columnName, columnType }) => {
      expect(columnType).toBe(
        floorSeedType("columnConfig", String(columnName)),
      );
    });
  });

  it("does not re-insert a column whose header drifted but whose column ID is intact", () => {
    const { grid } = floorFixture({
      sheetConfigHeaders: { sheetTitle: "Tab name" },
    });
    const { report } = applyFloor();

    expect(
      rowValues(grid, sheetConfigGid, sheetLayout.tableHeaderRowIndex),
    ).toEqual(defaultSheetConfigHeaders);
    expect(report).not.toContain("Recreated columns:");
    expect(report).toContain("Restored headers:");
  });

  it.each([
    {
      column: sc.sheetGid.header,
      options: { sheetConfigColumnOrder: ["sheetTitle", "letApiAccess"] },
    },
    {
      column: sc.letApiAccess.header,
      options: { sheetConfigColumnOrder: ["sheetGid", "sheetTitle"] },
    },
    {
      column: cc.columnId.header,
      options: {
        columnConfigColumnOrder: [
          "sheetGid",
          "sheetTitle",
          "header",
          "emptyValueAllowed",
        ],
      },
    },
    {
      column: cc.emptyValueAllowed.header,
      options: {
        columnConfigColumnOrder: [
          "sheetGid",
          "columnId",
          "sheetTitle",
          "header",
        ],
      },
    },
    {
      column: ssc.tableMenuSpace.header,
      options: {
        spreadsheetConfigColumnOrder: sscColumns.filter(
          (columnName) => columnName !== "tableMenuSpace",
        ),
      },
    },
    {
      column: ssc.tableMenuSpace.header,
      options: {
        spreadsheetConfigColumnOrder: [
          "fillRowIdsTimeLastRan",
          "tableMenuSpace",
          ...sscColumns.filter(
            (columnName) =>
              columnName !== "tableMenuSpace" &&
              columnName !== "fillRowIdsTimeLastRan",
          ),
        ],
      },
    },
  ] as const)(
    "throws naming $column when it can't be recreated, and sends no batch update",
    ({ column, options }) => {
      const { batchUpdateCount } = floorFixture(options);

      expect(() => applyFloor()).toThrow(`"${column}"`);
      expect(batchUpdateCount()).toBe(0);
    },
  );

  it("makes two grid reads and sends no batch update from inside ensure when nothing is missing", () => {
    const { batchUpdateCount, getByDataFilterCalls } = floorFixture();
    ConfigSheetFloor.init().ensure();

    expect(getByDataFilterCalls).toHaveLength(2);
    expect(batchUpdateCount()).toBe(0);
  });

  it("skips a recreated column the refetch still lacks in the label, data-value, column-type and edit-warning steps, without throwing, in two batch updates", () => {
    const { batchUpdateCount } = floorFixture({
      columnTypesAreUnset: true,
      columnConfigColumnOrder: [
        "sheetGid",
        "columnId",
        "sheetTitle",
        "emptyValueAllowed",
      ],
    });
    const { floor, report } = applyFloor();

    expect(batchUpdateCount()).toBe(2);
    expect(report).toContain("Set column types:");
    expect(report).not.toContain(`(${cc.header.columnId}) → TEXT`);
    expect(report).not.toContain("Restored headers:");
    expect(protectionsOf(floor, "columnConfig")).toHaveLength(1);
  });

  it.each([
    { sheetName: "sheetConfig", sheetGid: sheetConfigGid },
    { sheetName: "columnConfig", sheetGid: columnConfigGid },
  ] as const)(
    "creates a missing $sheetName tab at its GID with its seed title and Table, and reports it",
    ({ sheetName, sheetGid }) => {
      const { grid } = floorFixture({ omitSheetGids: [sheetGid] });
      const { report } = applyFloor();
      const seed = configSheetFloorSeed[sheetName];

      expect(grid.sheet(sheetGid).title).toBe(seed.title);
      expect(grid.sheet(sheetGid).tables.map((table) => table.name)).toEqual([
        seed.tableName,
      ]);
      expect(tableColumns(grid, sheetGid)).toEqual(
        seedTableColumns(seed.columns),
      );
      expect(report).toContain(`Created tabs: ${seed.title}`);
    },
  );

  it("places a created Table's header row, first column and one data row by the sheet layout", () => {
    const { grid } = floorFixture({
      omitSheetGids: [sheetConfigGid],
    });
    applyFloor();

    const headerRowIndex = sheetLayout.tableHeaderRowIndex;
    const startColIndex = sheetLayout.startTableColIndex;
    expect(grid.sheet(sheetConfigGid).tables[0]?.range).toEqual({
      startRowIndex: headerRowIndex,
      endRowIndex: headerRowIndex + 2,
      startColumnIndex: startColIndex,
      endColumnIndex:
        startColIndex + configSheetFloorSeed.sheetConfig.columns.length,
    });
  });

  it("creates a missing Spreadsheet Config at its GID with a Table carrying every seed column, the endpoint feedback columns included, and reports it", () => {
    const { grid } = floorFixture({
      omitSheetGids: [spreadsheetConfigGid],
    });
    const { report } = applyFloor();
    const seed = configSheetFloorSeed.spreadsheetConfig;

    expect(grid.sheet(spreadsheetConfigGid).title).toBe(seed.title);
    expect(
      grid.sheet(spreadsheetConfigGid).tables.map((table) => table.name),
    ).toEqual([seed.tableName]);
    expect(tableColumns(grid, spreadsheetConfigGid)).toEqual(
      seedTableColumns(floorSeedColumns("spreadsheetConfig")),
    );
    expect(
      tableColumns(grid, spreadsheetConfigGid).map(
        (column) => column.columnName,
      ),
    ).toEqual(
      expect.arrayContaining([
        ssc.fillRowIdsTimeLastRan.header,
        ssc.fillRowIdsRunStatus.header,
        ssc.syncConfigSheetRowsTimeLastRan.header,
        ssc.syncConfigSheetRowsRunStatus.header,
      ]),
    );
    expect(report).toContain(`Created tabs: ${seed.title}`);
  });

  it("creates a missing Spreadsheet Config with only Table menu space and the endpoint feedback columns", () => {
    const { grid } = floorFixture({
      omitSheetGids: [spreadsheetConfigGid],
    });
    applyFloor();

    expect(
      tableColumns(grid, spreadsheetConfigGid).map(
        (column) => column.columnName,
      ),
    ).toEqual(defaultSpreadsheetConfigHeaders);
  });

  it("leaves a created Spreadsheet Config's data row empty but for Not used in Table menu space", () => {
    const { grid } = floorFixture({
      omitSheetGids: [spreadsheetConfigGid],
    });
    applyFloor();

    expect(
      rowValues(grid, spreadsheetConfigGid, topDataRowIndex, sscColumns.length),
    ).toEqual(
      sscColumns.map((columnName) =>
        columnName === "tableMenuSpace" ? "Not used" : null,
      ),
    );
  });

  it("writes nothing into a created Spreadsheet Config's action row, and no data validation", () => {
    const { grid } = floorFixture({
      omitSheetGids: [spreadsheetConfigGid],
    });
    applyFloor();

    const sheet = grid.sheet(spreadsheetConfigGid);
    expect(
      rowValues(grid, spreadsheetConfigGid, actionRowIndex, sscColumns.length),
    ).toEqual(sscColumns.map(() => null));
    expect(
      sheet
        .rows()
        .flat()
        .filter(
          (cell) =>
            typeof cell === "object" &&
            cell?.dataValidationConditionType !== undefined,
        ),
    ).toEqual([]);
    expect(
      sheet.tables.flatMap((table) =>
        (table.columnProperties ?? []).filter(
          (column) => column.dataValidationRule !== undefined,
        ),
      ),
    ).toEqual([]);
  });

  it("creates a missing Value Config at its GID with a Table of one example column and two data rows, and reports it", () => {
    const { grid } = floorFixture({
      omitSheetGids: [valueConfigGid],
    });
    const { report } = applyFloor();
    const seed = configSheetFloorSeed.valueConfig;

    expect(grid.sheet(valueConfigGid).title).toBe(seed.title);
    const tables = grid.sheet(valueConfigGid).tables;
    expect(tables.map((table) => table.name)).toEqual(["valueConfig"]);
    expect(tableColumns(grid, valueConfigGid)).toEqual([
      { columnName: "Example value", columnType: "TEXT" },
    ]);
    const headerRowIndex = sheetLayout.tableHeaderRowIndex;
    const startColIndex = sheetLayout.startTableColIndex;
    expect(tables[0]?.range).toEqual({
      startRowIndex: headerRowIndex,
      endRowIndex: headerRowIndex + 3,
      startColumnIndex: startColIndex,
      endColumnIndex: startColIndex + 1,
    });
    expect(report).toContain(`Created tabs: ${seed.title}`);
  });

  it("fills a created Value Config's example column with a vcf column ID and its sample members, under the header its Table was created with", () => {
    const { grid } = floorFixture({
      omitSheetGids: [valueConfigGid],
    });
    applyFloor();

    const { seededValues } = configSheetFloorSeed.valueConfig.exampleColumn;
    const startColIndex = sheetLayout.startTableColIndex;
    expect(
      grid.sheet(valueConfigGid).values({
        startColumnIndex: startColIndex,
        endColumnIndex: startColIndex + 1,
      }),
    ).toEqual([
      [expect.stringMatching(/^c:vcf:/)],
      [null],
      [null],
      ["Example value"],
      ...seededValues.map((value) => [value]),
    ]);
  });

  it("gives an existing Value Config whose example column was deleted no example column back", () => {
    const { grid } = floorFixture();
    const before = sheetSnapshot(grid, valueConfigGid);
    applyFloor();

    expect(sheetSnapshot(grid, valueConfigGid)).toEqual(before);
  });

  it("gives a created Value Config no edit protection", () => {
    const { grid } = floorFixture({
      omitSheetGids: [valueConfigGid],
    });
    applyFloor();

    expect(grid.sheet(valueConfigGid).protectedRanges).toEqual([]);
  });

  it("names exactly Spreadsheet Config, Sheet Config and Column Config as the warned floor tabs", () => {
    expect(floorSheetNames()).toEqual([
      "spreadsheetConfig",
      "sheetConfig",
      "columnConfig",
    ]);
  });

  it("skips a created tab the refetch still lacks in the label, data-value and column-type steps, without throwing, in two batch updates", () => {
    const { batchUpdateCount } = floorFixture({
      columnTypesAreUnset: true,
      omitSheetGids: [sheetConfigGid],
      isDryRun: true,
    });
    const { report } = applyFloor();

    expect(batchUpdateCount()).toBe(2);
    expect(report).toContain("Set column types:");
    expect(report).not.toContain("Sheet Config ·");
  });

  it("never lets a recreatable column be one of a self-describing row's identity or declared columns", () => {
    floorSheetNames().forEach((sheetName) => {
      const selfDescribing: readonly string[] =
        selfDescribingRowColumns(sheetName);
      expect(
        floorRecreatableColumns(sheetName).filter((columnName) =>
          selfDescribing.includes(columnName),
        ),
      ).toEqual([]);
    });
    expect(selfDescribingRowColumns("columnConfig")).toContain("columnId");
  });
});
