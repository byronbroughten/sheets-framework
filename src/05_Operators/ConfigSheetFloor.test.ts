import { afterEach, beforeEach, describe, expect, it } from "vitest";

import type { ModelableEditProtection } from "../00_Source/RawSource/EditProtection";
import { installedConfigs } from "../01_SpreadsheetSchema/configRegister";
import {
  configSheetFloorSeed,
  floorSeedColumns,
} from "../01_SpreadsheetSchema/configSheetFloorSeed";
import { getSheetTraitByName } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import {
  clearSpreadsheetConfigOverlay,
  overlaySpreadsheetConfig,
  ssConfigGet,
} from "../01_SpreadsheetSchema/spreadsheetConfigTypes";
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
import { SpreadsheetConfigOperator } from "./SpreadsheetConfigOperator";

const { columnConfigs, spreadsheetConfig } = installedConfigs();

const spreadsheetConfigGid = getSheetTraitByName(
  "spreadsheetConfig",
  "sheetGid",
);
const sheetConfigGid = getSheetTraitByName("sheetConfig", "sheetGid");
const columnConfigGid = getSheetTraitByName("columnConfig", "sheetGid");
const valueConfigGid = getSheetTraitByName("valueConfig", "sheetGid");
const actionRowIndex = ssConfigGet("actionRowIndexBase0");
const topDataRowIndex = ssConfigGet("tableHeaderRowIndexBase0") + 1;
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
  "idDelimiter",
  "idHeader",
  "nameHeader",
  "startTableColumnIndexBase1",
  "columnIdRowIndexBase1",
  "columnGroupHeadingRowIndexBase1",
  "actionRowIndexBase1",
  "tableHeaderRowIndexBase1",
] as const;

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  clearSpreadsheetConfigOverlay();
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

function columnTypeUpdates(
  batchUpdateCalls: { requests?: GoogleAppsScript.Sheets.Schema.Request[] }[],
) {
  return (batchUpdateCalls[0]?.requests ?? []).flatMap((request) => {
    const table = request.updateTable?.table;
    if (table?.tableId === undefined) return [];
    return (table.columnProperties ?? []).flatMap((column) => {
      if (column.columnIndex === undefined || column.columnType === undefined) {
        return [];
      }
      return [
        {
          tableId: table.tableId,
          columnIndex: column.columnIndex,
          columnType: column.columnType,
        },
      ];
    });
  });
}

function firstFlushRequests(
  batchUpdateCalls: { requests?: GoogleAppsScript.Sheets.Schema.Request[] }[],
): GoogleAppsScript.Sheets.Schema.Request[] {
  return batchUpdateCalls[0]?.requests ?? [];
}

function sheetTitleUpdates(
  batchUpdateCalls: { requests?: GoogleAppsScript.Sheets.Schema.Request[] }[],
) {
  return firstFlushRequests(batchUpdateCalls).flatMap((request) => {
    const properties = request.updateSheetProperties?.properties;
    if (properties?.sheetId === undefined || properties.title === undefined) {
      return [];
    }
    return [{ sheetId: properties.sheetId, title: properties.title }];
  });
}

function tableNameUpdates(
  batchUpdateCalls: { requests?: GoogleAppsScript.Sheets.Schema.Request[] }[],
) {
  return firstFlushRequests(batchUpdateCalls).flatMap((request) => {
    const table = request.updateTable?.table;
    if (
      request.updateTable?.fields !== "name" ||
      table?.tableId === undefined ||
      table.name === undefined
    ) {
      return [];
    }
    return [{ tableId: table.tableId, name: table.name }];
  });
}

function cellUpdates(
  batchUpdateCalls: { requests?: GoogleAppsScript.Sheets.Schema.Request[] }[],
) {
  return firstFlushRequests(batchUpdateCalls).flatMap((request) => {
    const update = request.updateCells;
    const range = update?.range;
    if (
      range?.sheetId === undefined ||
      range.startRowIndex === undefined ||
      range.startColumnIndex === undefined
    ) {
      return [];
    }
    const userEnteredValue = update?.rows?.[0]?.values?.[0]?.userEnteredValue;
    return [
      {
        sheetId: range.sheetId,
        rowIndex: range.startRowIndex,
        colIndex: range.startColumnIndex,
        value: userEnteredValue?.stringValue ?? userEnteredValue?.numberValue,
      },
    ];
  });
}

function spreadsheetConfigCellUpdates(
  batchUpdateCalls: { requests?: GoogleAppsScript.Sheets.Schema.Request[] }[],
) {
  return cellUpdates(batchUpdateCalls).filter(
    (update) => update.sheetId === spreadsheetConfigGid,
  );
}

function columnInserts(
  batchUpdateCalls: { requests?: GoogleAppsScript.Sheets.Schema.Request[] }[],
) {
  return firstFlushRequests(batchUpdateCalls).flatMap((request) => {
    const range = request.insertDimension?.range;
    if (range?.dimension !== "COLUMNS") return [];
    return [{ sheetId: range.sheetId, startIndex: range.startIndex }];
  });
}

function addSheetRequests(
  batchUpdateCalls: { requests?: GoogleAppsScript.Sheets.Schema.Request[] }[],
) {
  return firstFlushRequests(batchUpdateCalls).flatMap((request) => {
    const properties = request.addSheet?.properties;
    return properties === undefined ? [] : [properties];
  });
}

// The adapter sends a Table's columns on the updateTable after its addTable, not on the addTable.
function addTableRequests(
  batchUpdateCalls: { requests?: GoogleAppsScript.Sheets.Schema.Request[] }[],
) {
  const requests = firstFlushRequests(batchUpdateCalls);
  return requests.flatMap((request) => {
    const table = request.addTable?.table;
    if (table === undefined) return [];
    const columnProperties = requests.find(
      (candidate) => candidate.updateTable?.table?.tableId === table.tableId,
    )?.updateTable?.table?.columnProperties;
    return [{ ...table, columnProperties }];
  });
}

function requestsOnSheet(
  requests: GoogleAppsScript.Sheets.Schema.Request[],
  sheetId: number,
): GoogleAppsScript.Sheets.Schema.Request[] {
  return requests.filter((request) =>
    JSON.stringify(request).includes(`"sheetId":${sheetId}`),
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

const sheetConfigColumns = ["sheetGid", "sheetTitle", "letApiAccess"] as const;
const columnConfigColumns = [
  "sheetGid",
  "columnId",
  "sheetTitle",
  "header",
  "emptyValueAllowed",
] as const;

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
    spreadsheetConfigLayoutValues?: Partial<
      Record<(typeof sscColumns)[number], string | number>
    >;
    startTableColIndex?: number;
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
  const startTableColIndex = options.startTableColIndex ?? 0;
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
  const dataRow = sscOrder.map((columnName) => {
    const layoutValue = options.spreadsheetConfigLayoutValues?.[columnName];
    if (layoutValue !== undefined) return layoutValue;
    if (columnName === "tableMenuSpace") {
      return options.tableMenuSpaceValue ?? "Not used";
    }
    if (columnName === "idDelimiter") return ":";
    if (columnName === "idHeader") return "ID";
    if (columnName === "nameHeader") return "Name";
    if (columnName === "startTableColumnIndexBase1") return 1;
    if (columnName === "columnIdRowIndexBase1") return 1;
    if (columnName === "columnGroupHeadingRowIndexBase1") return 2;
    if (columnName === "actionRowIndexBase1") return 3;
    if (columnName === "tableHeaderRowIndexBase1") return 4;
    return "";
  });

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

function addedProtectedRanges(
  batchUpdateCalls: { requests?: GoogleAppsScript.Sheets.Schema.Request[] }[],
) {
  return (batchUpdateCalls[0]?.requests ?? []).flatMap((request) => {
    const protection = request.addProtectedRange?.protectedRange;
    return protection === undefined ? [] : [protection];
  });
}

function deletedProtectionIds(
  batchUpdateCalls: { requests?: GoogleAppsScript.Sheets.Schema.Request[] }[],
) {
  return (batchUpdateCalls[0]?.requests ?? []).flatMap((request) => {
    const id = request.deleteProtectedRange?.protectedRangeId;
    return id === undefined ? [] : [id];
  });
}

const spreadsheetConfigLayoutRange = {
  sheetId: spreadsheetConfigGid,
  startRowIndex: topDataRowIndex,
  startColumnIndex: 5,
  endColumnIndex: 13,
};
const spreadsheetConfigSelectorRanges = [
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
const spreadsheetConfigEditableRanges = [
  ...spreadsheetConfigSelectorRanges,
  spreadsheetConfigLayoutRange,
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
    const { batchUpdateCalls } = floorFixture();
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
    expect(
      addedProtectedRanges(batchUpdateCalls).map(
        (protection) => protection.range?.sheetId,
      ),
    ).not.toContain(valueConfigGid);
  });

  it("queues nothing on a second run", () => {
    const { batchUpdateCalls } = floorFixture();
    const { floor } = applyFloor();
    floor.ensure();
    floor.ss.batchUpdateGSheets();
    expect(batchUpdateCalls).toHaveLength(1);
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

  it("queues no protection change on a second sync against the fixture the first one left", () => {
    const { batchUpdateCalls } = floorFixture();
    applyFloor();
    expect(batchUpdateCalls).toHaveLength(1);

    const second = applyFloor();

    expect(second.report).toBe("");
    expect(batchUpdateCalls).toHaveLength(1);
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
      ...spreadsheetConfigSelectorRanges,
      { ...spreadsheetConfigLayoutRange, endColumnIndex: 14 },
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
    const { batchUpdateCalls } = floorFixture({
      sheetConfigProtections: [
        {
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
        },
      ],
    });
    const { floor } = applyFloor();

    expect(deletedProtectionIds(batchUpdateCalls)).not.toContain(99);
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
    const { batchUpdateCalls } = floorFixture({
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

    expect(deletedProtectionIds(batchUpdateCalls)).toContain(41);
    expect(protectionsOf(floor, "spreadsheetConfig")).toHaveLength(1);
    expect(protectionsOf(floor, "spreadsheetConfig")[0]).toMatchObject({
      description: floorWarningDescription("Spreadsheet Config"),
      range: { sheetId: spreadsheetConfigGid },
    });
  });

  it("sets a floor column whose type differs from the seed and reports it", () => {
    const { batchUpdateCalls } = floorFixture({
      spreadsheetConfigColumnTypes: { tableMenuSpace: "DOUBLE" },
    });
    const { report } = applyFloor();

    expect(report).toContain("Set column types:");
    expect(report).toContain(
      `Spreadsheet Config · Table menu space (${ssc.tableMenuSpace.columnId}) → TEXT`,
    );
    expect(columnTypeUpdates(batchUpdateCalls)).toContainEqual({
      tableId: `fake-table-${spreadsheetConfigGid}`,
      columnIndex: 0,
      columnType: "TEXT",
    });
    expect(
      (batchUpdateCalls[0]?.requests ?? []).filter(
        (request) => request.updateTable !== undefined,
      ),
    ).toHaveLength(1);
  });

  it("leaves matching floor column types unchanged", () => {
    const { batchUpdateCalls } = floorFixture();
    const { report } = applyFloor();

    expect(report).not.toContain("Set column types:");
    expect(columnTypeUpdates(batchUpdateCalls)).toEqual([]);
  });

  it("does not set a type on a column added to Column Config", () => {
    const { batchUpdateCalls } = floorFixture({
      columnTypesAreUnset: true,
      extraColumnConfigColumn: addedColumnConfigColumn,
    });
    applyFloor();

    expect(columnTypeUpdates(batchUpdateCalls)).not.toContainEqual(
      expect.objectContaining({
        tableId: `fake-table-${columnConfigGid}`,
        columnIndex: 5,
      }),
    );
  });

  it("does not queue column-type updates on a second run after setting them", () => {
    const { batchUpdateCalls } = floorFixture({ columnTypesAreUnset: true });
    applyFloor();

    expect(columnTypeUpdates(batchUpdateCalls).length).toBeGreaterThan(0);
    applyFloor();
    expect(batchUpdateCalls).toHaveLength(1);
  });

  it("sets a drifted floor column type when the Table starts after column A", () => {
    overlaySpreadsheetConfig({
      ...spreadsheetConfig,
      startTableColIndexBase0: 1,
    });
    const { batchUpdateCalls } = floorFixture({
      startTableColIndex: 1,
      spreadsheetConfigColumnTypes: { tableMenuSpace: "DOUBLE" },
    });
    applyFloor();

    expect(columnTypeUpdates(batchUpdateCalls)).toContainEqual({
      tableId: `fake-table-${spreadsheetConfigGid}`,
      columnIndex: 0,
      columnType: "TEXT",
    });
    applyFloor();
    expect(batchUpdateCalls).toHaveLength(1);
  });

  it("renames a drifted floor tab back and reports it", () => {
    const { batchUpdateCalls } = floorFixture({
      spreadsheetConfigTitle: "Old Spreadsheet Config",
    });
    const { report } = applyFloor();

    expect(report).toContain(
      'Restored tab titles: "Old Spreadsheet Config" → Spreadsheet Config',
    );
    expect(sheetTitleUpdates(batchUpdateCalls)).toContainEqual({
      sheetId: spreadsheetConfigGid,
      title: "Spreadsheet Config",
    });
  });

  it("names a restored tab by its seed title in the rest of the report", () => {
    floorFixture({
      spreadsheetConfigTitle: "Old Spreadsheet Config",
      spreadsheetConfigHeaders: { tableMenuSpace: "Menu spacer" },
      spreadsheetConfigColumnTypes: { idHeader: "DOUBLE" },
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
      `Spreadsheet Config · ID header (${ssc.idHeader.columnId}) → TEXT`,
    );
    expect(report).toContain(
      "Covered added columns: Spreadsheet Config · Notes",
    );
    expect(report).not.toContain("Old Spreadsheet Config ·");
  });

  it("renames Value Config's tab back and reports it", () => {
    const { batchUpdateCalls } = floorFixture({
      valueConfig: { title: "Values" },
    });
    const { report } = applyFloor();

    expect(report).toContain('Restored tab titles: "Values" → Value Config');
    expect(sheetTitleUpdates(batchUpdateCalls)).toContainEqual({
      sheetId: valueConfigGid,
      title: "Value Config",
    });
  });

  it("returns the floor notice for a renamed Value Config off freshly fetched sheet properties and sends no batch update", () => {
    const { batchUpdateCalls, getCalls } = floorFixture({
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
    expect(batchUpdateCalls).toHaveLength(0);
  });

  it("renames a wrongly named floor Table and reports it", () => {
    const { batchUpdateCalls } = floorFixture({
      spreadsheetConfigTableName: "wrongTable",
    });
    const { report } = applyFloor();

    expect(report).toContain(
      `Restored Table names: Spreadsheet Config's Table "wrongTable" → spreadsheetConfig`,
    );
    expect(tableNameUpdates(batchUpdateCalls)).toContainEqual({
      tableId: `fake-table-${spreadsheetConfigGid}`,
      name: "spreadsheetConfig",
    });
  });

  it("renames Value Config's Table back to valueConfig", () => {
    const { batchUpdateCalls } = floorFixture({
      valueConfig: { tableName: "values" },
    });
    const { report } = applyFloor();

    expect(report).toContain(
      `Restored Table names: Value Config's Table "values" → valueConfig`,
    );
    expect(tableNameUpdates(batchUpdateCalls)).toContainEqual({
      tableId: `fake-table-${valueConfigGid}`,
      name: "valueConfig",
    });
  });

  it("throws naming the tab when a floor title sits on the wrong GID, and flushes nothing", () => {
    const { batchUpdateCalls } = floorFixture({
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
    expect(batchUpdateCalls).toHaveLength(0);
  });

  it("throws naming the tab when a floor tab has no Table, and flushes nothing", () => {
    const { batchUpdateCalls } = floorFixture({
      omitSpreadsheetConfigTable: true,
    });

    expect(() => applyFloor()).toThrow(
      'Floor tab "Spreadsheet Config" has no Table.',
    );
    expect(batchUpdateCalls).toHaveLength(0);
  });

  it("throws naming the tab when several Tables are present and none has the floor name, and flushes nothing", () => {
    const { batchUpdateCalls } = floorFixture({
      spreadsheetConfigTableName: "firstWrong",
      spreadsheetConfigExtraTables: [{ endRowIndex: 5, name: "secondWrong" }],
    });

    expect(() => applyFloor()).toThrow(
      'Floor tab "Spreadsheet Config" has several Tables and none is named spreadsheetConfig.',
    );
    expect(batchUpdateCalls).toHaveLength(0);
  });

  it("throws naming the tab when several Tables are present and one has the floor name, and flushes nothing", () => {
    const { batchUpdateCalls } = floorFixture({
      spreadsheetConfigExtraTables: [{ endRowIndex: 5, name: "extra" }],
    });

    expect(() => applyFloor()).toThrow(
      '1 sheet(s) have more than one Table — delete the extras so each sheet has exactly one: "Spreadsheet Config"',
    );
    expect(batchUpdateCalls).toHaveLength(0);
  });

  it("overwrites a drifted floor header and reports it", () => {
    const { batchUpdateCalls } = floorFixture({
      spreadsheetConfigHeaders: { tableMenuSpace: "Menu spacer" },
    });
    const { report } = applyFloor();

    expect(report).toContain("Restored headers:");
    expect(report).toContain(
      `Spreadsheet Config · Menu spacer (${ssc.tableMenuSpace.columnId}) → Table menu space`,
    );
    expect(cellUpdates(batchUpdateCalls)).toContainEqual({
      sheetId: spreadsheetConfigGid,
      rowIndex: 3,
      colIndex: 0,
      value: "Table menu space",
    });
  });

  it("overwrites a drifted floor column ID and reports it", () => {
    const { batchUpdateCalls } = floorFixture({
      spreadsheetConfigColumnIds: { tableMenuSpace: "c:sscf:drifted" },
    });
    const { report } = applyFloor();

    expect(report).toContain("Restored column IDs:");
    expect(report).toContain(
      `Spreadsheet Config · Table menu space (c:sscf:drifted) → ${ssc.tableMenuSpace.columnId}`,
    );
    expect(cellUpdates(batchUpdateCalls)).toContainEqual({
      sheetId: spreadsheetConfigGid,
      rowIndex: 0,
      colIndex: 0,
      value: ssc.tableMenuSpace.columnId,
    });
  });

  it("overwrites a drifted Sheet GID column ID on Sheet Config without throwing, and reports it", () => {
    const { batchUpdateCalls } = floorFixture({
      sheetConfigColumnIds: { sheetGid: "c:scf:drifted" },
    });
    const { report } = applyFloor();

    expect(report).toContain(
      `Sheet Config · ${sc.sheetGid.header} (c:scf:drifted) → ${sc.sheetGid.columnId}`,
    );
    expect(cellUpdates(batchUpdateCalls)).toContainEqual({
      sheetId: sheetConfigGid,
      rowIndex: 0,
      colIndex: 0,
      value: sc.sheetGid.columnId,
    });
  });

  it("overwrites a drifted Column ID column ID on Column Config without throwing, and reports it", () => {
    const { batchUpdateCalls } = floorFixture({
      columnConfigColumnIds: { columnId: "c:ccf:drifted" },
    });
    const { report } = applyFloor();

    expect(report).toContain(
      `Column Config · ${cc.columnId.header} (c:ccf:drifted) → ${cc.columnId.columnId}`,
    );
    expect(cellUpdates(batchUpdateCalls)).toContainEqual({
      sheetId: columnConfigGid,
      rowIndex: 0,
      colIndex: 1,
      value: cc.columnId.columnId,
    });
  });

  it("carves nothing on Sheet Config the sync its Sheet GID header has drifted, and restores the header without throwing", () => {
    const { batchUpdateCalls } = floorFixture({
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
    expect(cellUpdates(batchUpdateCalls)).toContainEqual({
      sheetId: sheetConfigGid,
      rowIndex: 3,
      colIndex: 0,
      value: sc.sheetGid.header,
    });
  });

  it("overwrites a drifted floor group heading and reports it", () => {
    const { batchUpdateCalls } = floorFixture({
      spreadsheetConfigGroupHeadings: { idHeader: "Rules" },
    });
    const { report } = applyFloor();

    expect(report).toContain("Restored group headings:");
    expect(report).toContain(
      `Spreadsheet Config · ID header (${ssc.idHeader.columnId}) → Spreadsheet Rules`,
    );
    expect(cellUpdates(batchUpdateCalls)).toContainEqual({
      sheetId: spreadsheetConfigGid,
      rowIndex: 1,
      colIndex: sscColumns.indexOf("idHeader"),
      value: "Spreadsheet Rules",
    });
  });

  it("leaves an extra non-floor column's header, column ID and group heading alone", () => {
    const { batchUpdateCalls } = floorFixture({
      extraSpreadsheetConfigColumn: {
        columnId: "c:sscf:notes",
        header: "Notes",
        groupHeading: "Mine",
      },
    });
    applyFloor();

    const extraColIndex = sscColumns.length;
    expect(cellUpdates(batchUpdateCalls)).not.toContainEqual(
      expect.objectContaining({
        sheetId: spreadsheetConfigGid,
        colIndex: extraColIndex,
      }),
    );
  });

  it("writes Not used into a blank Table menu space data cell and reports it", () => {
    const { batchUpdateCalls } = floorFixture({ tableMenuSpaceValue: "" });
    const { report } = applyFloor();

    expect(report).toContain('Restored Table menu space: "" → Not used');
    const dataCellUpdates = cellUpdates(batchUpdateCalls).filter(
      (update) => update.rowIndex === topDataRowIndex,
    );
    expect(dataCellUpdates).toEqual([
      {
        sheetId: spreadsheetConfigGid,
        rowIndex: topDataRowIndex,
        colIndex: 0,
        value: "Not used",
      },
    ]);
  });

  it("queues no Table menu space update when the cell already reads Not used", () => {
    const { batchUpdateCalls } = floorFixture();
    const { report } = applyFloor();

    expect(report).not.toContain("Table menu space");
    expect(
      cellUpdates(batchUpdateCalls).filter(
        (update) => update.rowIndex === topDataRowIndex,
      ),
    ).toEqual([]);
  });

  it("overwrites an edited Table menu space data cell and reports it", () => {
    const { batchUpdateCalls } = floorFixture({ tableMenuSpaceValue: "notes" });
    const { report } = applyFloor();

    expect(report).toContain('Restored Table menu space: "notes" → Not used');
    expect(cellUpdates(batchUpdateCalls)).toContainEqual({
      sheetId: spreadsheetConfigGid,
      rowIndex: topDataRowIndex,
      colIndex: 0,
      value: "Not used",
    });
  });

  it("recreates a missing Column Config Header column at the Table end with its header, column ID and heading, and reports it", () => {
    const { batchUpdateCalls } = floorFixture({
      columnConfigColumnOrder: [
        "sheetGid",
        "columnId",
        "sheetTitle",
        "emptyValueAllowed",
      ],
    });
    const { report } = applyFloor();

    expect(columnInserts(batchUpdateCalls)).toEqual([
      { sheetId: columnConfigGid, startIndex: 4 },
    ]);
    const writes = cellUpdates(batchUpdateCalls).filter(
      (update) => update.sheetId === columnConfigGid && update.colIndex === 4,
    );
    expect(
      writes
        .map(({ rowIndex, value }) => ({ rowIndex, value }))
        .sort((left, right) => left.rowIndex - right.rowIndex),
    ).toEqual([
      { rowIndex: 0, value: cc.header.columnId },
      { rowIndex: 1, value: "" },
      { rowIndex: 3, value: cc.header.header },
    ]);
    expect(report).toContain(
      `Recreated columns: Column Config · ${cc.header.header} (${cc.header.columnId})`,
    );
  });

  it("recreates two missing columns on one tab as two Table-end inserts, in seed order", () => {
    const { batchUpdateCalls } = floorFixture({
      columnConfigColumnOrder: ["sheetGid", "columnId", "emptyValueAllowed"],
    });
    applyFloor();

    expect(columnInserts(batchUpdateCalls)).toEqual([
      { sheetId: columnConfigGid, startIndex: 3 },
      { sheetId: columnConfigGid, startIndex: 4 },
    ]);
    const headerWrites = cellUpdates(batchUpdateCalls).filter(
      (update) => update.sheetId === columnConfigGid && update.rowIndex === 3,
    );
    expect(headerWrites).toEqual([
      expect.objectContaining({ colIndex: 3, value: cc.sheetTitle.header }),
      expect.objectContaining({ colIndex: 4, value: cc.header.header }),
    ]);
  });

  it("sends no column-type update for a sheet in the flush that inserts a column on it", () => {
    const { batchUpdateCalls } = floorFixture({
      columnTypesAreUnset: true,
      columnConfigColumnOrder: [
        "sheetGid",
        "columnId",
        "sheetTitle",
        "emptyValueAllowed",
      ],
    });
    applyFloor();

    expect(columnInserts(batchUpdateCalls)).toEqual([
      { sheetId: columnConfigGid, startIndex: 4 },
    ]);
    expect(columnTypeUpdates(batchUpdateCalls)).not.toContainEqual(
      expect.objectContaining({ tableId: `fake-table-${columnConfigGid}` }),
    );
  });

  it("does not re-insert a column whose header drifted but whose column ID is intact", () => {
    const { batchUpdateCalls } = floorFixture({
      sheetConfigHeaders: { sheetTitle: "Tab name" },
    });
    const { report } = applyFloor();

    expect(columnInserts(batchUpdateCalls)).toEqual([]);
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
      column: ssc.idHeader.header,
      options: {
        spreadsheetConfigColumnOrder: sscColumns.filter(
          (columnName) => columnName !== "idHeader",
        ),
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
    "throws naming $column when it can't be recreated, and flushes nothing",
    ({ column, options }) => {
      const { batchUpdateCalls } = floorFixture(options);

      expect(() => applyFloor()).toThrow(`"${column}"`);
      expect(batchUpdateCalls).toHaveLength(0);
    },
  );

  it("makes two grid reads and sends no batch update from inside ensure when nothing is missing", () => {
    const { batchUpdateCalls, getByDataFilterCalls } = floorFixture();
    ConfigSheetFloor.init().ensure();

    expect(getByDataFilterCalls).toHaveLength(2);
    expect(batchUpdateCalls).toHaveLength(0);
  });

  it("skips a recreated column the refetch still lacks in the label, data-value, column-type and edit-warning steps, without throwing", () => {
    const { batchUpdateCalls } = floorFixture({
      columnTypesAreUnset: true,
      columnConfigColumnOrder: [
        "sheetGid",
        "columnId",
        "sheetTitle",
        "emptyValueAllowed",
      ],
    });
    const { floor, report } = applyFloor();

    expect(batchUpdateCalls).toHaveLength(2);
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
      const { batchUpdateCalls } = floorFixture({ omitSheetGids: [sheetGid] });
      const { report } = applyFloor();
      const seed = configSheetFloorSeed[sheetName];

      expect(addSheetRequests(batchUpdateCalls)).toEqual([
        expect.objectContaining({ sheetId: sheetGid, title: seed.title }),
      ]);
      const tables = addTableRequests(batchUpdateCalls);
      expect(tables.map((table) => table.name)).toEqual([seed.tableName]);
      expect(tables[0]?.columnProperties).toEqual(
        seed.columns.map((column, columnIndex) => ({
          columnIndex,
          columnName: column.header,
          columnType: column.columnType,
        })),
      );
      expect(report).toContain(`Created tabs: ${seed.title}`);
    },
  );

  it("places a created Table's header row, first column and one data row by the generated layout", () => {
    const { batchUpdateCalls } = floorFixture({
      omitSheetGids: [sheetConfigGid],
    });
    applyFloor();

    const headerRowIndex = spreadsheetConfig.tableHeaderRowIndexBase0;
    const startColIndex = spreadsheetConfig.startTableColIndexBase0;
    const table = addTableRequests(batchUpdateCalls)[0];
    expect(
      table?.columnProperties?.map((column) => column.columnIndex),
    ).toEqual([0, 1, 2]);
    expect(table?.range).toEqual({
      sheetId: sheetConfigGid,
      startRowIndex: headerRowIndex,
      endRowIndex: headerRowIndex + 2,
      startColumnIndex: startColIndex,
      endColumnIndex:
        startColIndex + configSheetFloorSeed.sheetConfig.columns.length,
    });
  });

  it("creates a missing Spreadsheet Config at its GID with a Table carrying every seed column, the endpoint feedback columns included, and reports it", () => {
    const { batchUpdateCalls } = floorFixture({
      omitSheetGids: [spreadsheetConfigGid],
    });
    const { report } = applyFloor();
    const seed = configSheetFloorSeed.spreadsheetConfig;

    expect(addSheetRequests(batchUpdateCalls)).toEqual([
      expect.objectContaining({
        sheetId: spreadsheetConfigGid,
        title: seed.title,
      }),
    ]);
    const tables = addTableRequests(batchUpdateCalls);
    expect(tables.map((table) => table.name)).toEqual([seed.tableName]);
    expect(tables[0]?.columnProperties).toEqual(
      floorSeedColumns("spreadsheetConfig").map((column, columnIndex) => ({
        columnIndex,
        columnName: column.header,
        columnType: column.columnType,
      })),
    );
    expect(
      tables[0]?.columnProperties?.map((column) => column.columnName),
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

  it("seeds a created Spreadsheet Config's data row with the generated layout values in base 1", () => {
    const { batchUpdateCalls } = floorFixture({
      omitSheetGids: [spreadsheetConfigGid],
    });
    applyFloor();

    const table = addTableRequests(batchUpdateCalls)[0];
    const colIndexOf = (header: string) =>
      (table?.range?.startColumnIndex ?? 0) +
      (table?.columnProperties?.find((column) => column.columnName === header)
        ?.columnIndex ?? -1);
    const dataRowIndex = (table?.range?.startRowIndex ?? 0) + 1;
    const expected = [
      { header: ssc.idHeader.header, value: spreadsheetConfig.idHeader },
      { header: ssc.nameHeader.header, value: spreadsheetConfig.nameHeader },
      { header: ssc.idDelimiter.header, value: spreadsheetConfig.idDelimiter },
      {
        header: ssc.startTableColumnIndexBase1.header,
        value: spreadsheetConfig.startTableColIndexBase0 + 1,
      },
      {
        header: ssc.columnIdRowIndexBase1.header,
        value: spreadsheetConfig.columnIdRowIdxBase0 + 1,
      },
      {
        header: ssc.columnGroupHeadingRowIndexBase1.header,
        value: spreadsheetConfig.columnGroupHeadingRowIndexBase0 + 1,
      },
      {
        header: ssc.actionRowIndexBase1.header,
        value: spreadsheetConfig.actionRowIndexBase0 + 1,
      },
      {
        header: ssc.tableHeaderRowIndexBase1.header,
        value: spreadsheetConfig.tableHeaderRowIndexBase0 + 1,
      },
    ].map(({ header, value }) => ({
      sheetId: spreadsheetConfigGid,
      rowIndex: dataRowIndex,
      colIndex: colIndexOf(header),
      value,
    }));

    expect(spreadsheetConfigCellUpdates(batchUpdateCalls)).toEqual(expected);
  });

  it("fetches a created Spreadsheet Config back as the generated spreadsheetConfig", () => {
    const { batchUpdateCalls } = floorFixture({
      omitSheetGids: [spreadsheetConfigGid],
    });
    applyFloor();

    const table = addTableRequests(batchUpdateCalls)[0];
    const startColIndex = table?.range?.startColumnIndex ?? 0;
    const headers = (table?.columnProperties ?? []).map(
      (column) => column.columnName ?? null,
    );
    // The floor restores a created tab's column IDs on a later pass.
    const columnIdByHeader = new Map<string | null, string>(
      Object.values(columnConfigs.spreadsheetConfig).map((column) => [
        column.header,
        column.columnId,
      ]),
    );
    const rowsByIndex: Record<number, FakeCell[]> = {
      [spreadsheetConfig.columnIdRowIdxBase0]: headers.map(
        (header) => columnIdByHeader.get(header) ?? null,
      ),
      [table?.range?.startRowIndex ?? 0]: headers,
    };
    spreadsheetConfigCellUpdates(batchUpdateCalls).forEach(
      ({ rowIndex, colIndex, value }) => {
        const row = (rowsByIndex[rowIndex] ??= []);
        row[colIndex - startColIndex] = value ?? null;
      },
    );
    stubSheetsService({
      sheets: [
        {
          sheetId: spreadsheetConfigGid,
          title: configSheetFloorSeed.spreadsheetConfig.title,
          rows: buildGridRows(rowsByIndex),
          table: { endRowIndex: table?.range?.endRowIndex ?? 0 },
        },
      ],
    });

    expect(SpreadsheetConfigOperator.init().fetchLiveConfig()).toEqual(
      spreadsheetConfig,
    );
  });

  it("sends a created Spreadsheet Config's seeded values in the same batch as, and after, its add-sheet and add-Table", () => {
    const { batchUpdateCalls } = floorFixture({
      omitSheetGids: [spreadsheetConfigGid],
    });
    applyFloor();

    const requests = firstFlushRequests(batchUpdateCalls);
    const addSheetAt = requests.findIndex(
      (request) =>
        request.addSheet?.properties?.sheetId === spreadsheetConfigGid,
    );
    const addTableAt = requests.findIndex(
      (request) =>
        request.addTable?.table?.name ===
        configSheetFloorSeed.spreadsheetConfig.tableName,
    );
    const seedAts = requests.flatMap((request, index) =>
      request.updateCells?.range?.sheetId === spreadsheetConfigGid
        ? [index]
        : [],
    );
    expect(addSheetAt).toBeGreaterThanOrEqual(0);
    expect(addTableAt).toBeGreaterThan(addSheetAt);
    expect(seedAts).toHaveLength(8);
    expect(Math.min(...seedAts)).toBeGreaterThan(addTableAt);
  });

  it("writes nothing into a created Spreadsheet Config's Table menu space data cell", () => {
    const { batchUpdateCalls } = floorFixture({
      omitSheetGids: [spreadsheetConfigGid],
    });
    applyFloor();

    const tableMenuSpaceColIndex =
      addTableRequests(batchUpdateCalls)[0]?.range?.startColumnIndex;
    expect(
      spreadsheetConfigCellUpdates(batchUpdateCalls).filter(
        (cell) => cell.colIndex === tableMenuSpaceColIndex,
      ),
    ).toEqual([]);
  });

  it("writes nothing into a created Spreadsheet Config's action row, and no data validation", () => {
    const { batchUpdateCalls } = floorFixture({
      omitSheetGids: [spreadsheetConfigGid],
    });
    applyFloor();

    const requests = batchUpdateCalls.flatMap((call) => call.requests ?? []);
    expect(
      spreadsheetConfigCellUpdates(batchUpdateCalls).filter(
        (cell) => cell.rowIndex === actionRowIndex,
      ),
    ).toEqual([]);
    expect(requests.filter((request) => request.setDataValidation)).toEqual([]);
  });

  it("leaves an existing Spreadsheet Config's changed layout values alone", () => {
    const { batchUpdateCalls } = floorFixture({
      spreadsheetConfigLayoutValues: {
        idHeader: "Key",
        idDelimiter: "-",
        startTableColumnIndexBase1: 2,
      },
    });
    const { report } = applyFloor();

    expect(addSheetRequests(batchUpdateCalls)).toEqual([]);
    expect(
      spreadsheetConfigCellUpdates(batchUpdateCalls).filter(
        (cell) => cell.rowIndex === topDataRowIndex,
      ),
    ).toEqual([]);
    expect(report).not.toContain("Created tabs");
  });

  it("creates a missing Value Config at its GID with a Table of one example column and two data rows, and reports it", () => {
    const { batchUpdateCalls } = floorFixture({
      omitSheetGids: [valueConfigGid],
    });
    const { report } = applyFloor();
    const seed = configSheetFloorSeed.valueConfig;

    expect(addSheetRequests(batchUpdateCalls)).toEqual([
      expect.objectContaining({ sheetId: valueConfigGid, title: seed.title }),
    ]);
    const tables = addTableRequests(batchUpdateCalls);
    expect(tables.map((table) => table.name)).toEqual(["valueConfig"]);
    expect(tables[0]?.columnProperties).toEqual([
      { columnIndex: 0, columnName: "Example value", columnType: "TEXT" },
    ]);
    const headerRowIndex = spreadsheetConfig.tableHeaderRowIndexBase0;
    const startColIndex = spreadsheetConfig.startTableColIndexBase0;
    expect(tables[0]?.range).toEqual({
      sheetId: valueConfigGid,
      startRowIndex: headerRowIndex,
      endRowIndex: headerRowIndex + 3,
      startColumnIndex: startColIndex,
      endColumnIndex: startColIndex + 1,
    });
    expect(report).toContain(`Created tabs: ${seed.title}`);
  });

  it("sends a created Value Config's sample members and a vcf column ID, and no header write, in the same batch as and after its add-sheet and add-Table", () => {
    const { batchUpdateCalls } = floorFixture({
      omitSheetGids: [valueConfigGid],
    });
    applyFloor();

    const requests = firstFlushRequests(batchUpdateCalls);
    const addSheetAt = requests.findIndex(
      (request) => request.addSheet?.properties?.sheetId === valueConfigGid,
    );
    const addTableAt = requests.findIndex(
      (request) => request.addTable?.table?.name === "valueConfig",
    );
    const seedAts = requests.flatMap((request, index) =>
      request.updateCells?.range?.sheetId === valueConfigGid ? [index] : [],
    );
    expect(addSheetAt).toBeGreaterThanOrEqual(0);
    expect(addTableAt).toBeGreaterThan(addSheetAt);
    expect(seedAts).toHaveLength(3);
    expect(Math.min(...seedAts)).toBeGreaterThan(addTableAt);

    const startColIndex = spreadsheetConfig.startTableColIndexBase0;
    const valueConfigCells = cellUpdates(batchUpdateCalls).filter(
      (cell) => cell.sheetId === valueConfigGid,
    );
    expect(valueConfigCells).toEqual([
      {
        sheetId: valueConfigGid,
        rowIndex: spreadsheetConfig.columnIdRowIdxBase0,
        colIndex: startColIndex,
        value: expect.stringMatching(/^c:vcf:/),
      },
      ...configSheetFloorSeed.valueConfig.exampleColumn.seededValues.map(
        (value, memberIndex) => ({
          sheetId: valueConfigGid,
          rowIndex: topDataRowIndex + memberIndex,
          colIndex: startColIndex,
          value,
        }),
      ),
    ]);
    expect(
      valueConfigCells.filter(
        (cell) => cell.rowIndex === spreadsheetConfig.tableHeaderRowIndexBase0,
      ),
    ).toEqual([]);
  });

  it("gives an existing Value Config whose example column was deleted no example column back", () => {
    const { batchUpdateCalls } = floorFixture();
    applyFloor();

    const requests = batchUpdateCalls.flatMap((call) => call.requests ?? []);
    expect(requestsOnSheet(requests, valueConfigGid)).toEqual([]);
  });

  it("gives a created Value Config no edit-protection request", () => {
    const { batchUpdateCalls } = floorFixture({
      omitSheetGids: [valueConfigGid],
    });
    applyFloor();

    expect(
      addedProtectedRanges(batchUpdateCalls).map(
        (protection) => protection.range?.sheetId,
      ),
    ).not.toContain(valueConfigGid);
  });

  it("names exactly Spreadsheet Config, Sheet Config and Column Config as the warned floor tabs", () => {
    expect(floorSheetNames()).toEqual([
      "spreadsheetConfig",
      "sheetConfig",
      "columnConfig",
    ]);
  });

  it("skips a created tab the refetch still lacks in the label, data-value, column-type and edit-warning steps, without throwing", () => {
    const { batchUpdateCalls } = floorFixture({
      columnTypesAreUnset: true,
      omitSheetGids: [sheetConfigGid],
    });
    const { floor, report } = applyFloor();

    expect(batchUpdateCalls).toHaveLength(2);
    const laterRequests = batchUpdateCalls[1]?.requests ?? [];
    expect(requestsOnSheet(laterRequests, sheetConfigGid)).toEqual([]);
    expect(report).toContain("Set column types:");
    expect(report).not.toContain("Sheet Config ·");
    expect(protectionsOf(floor, "columnConfig")).toHaveLength(1);
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
