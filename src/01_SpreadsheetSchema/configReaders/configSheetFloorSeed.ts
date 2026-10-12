import { Obj } from "@byronbroughten/utils/obj";

import type { TableColumnType } from "../../00_Source/RawSource/RawSource";
import {
  getColumnTraitByName,
  getSheetColumnNames,
} from "./columnConfigsTypes";
import {
  type TableConfig,
  tableConfigAloneOnGid,
  tableConfigsByTableId,
} from "./tableConfigsTypes";

export type FloorColumnType = Extract<
  TableColumnType,
  "TEXT" | "DOUBLE" | "BOOLEAN"
>;

export interface FloorSeedColumn {
  header: string;
  columnType: FloorColumnType;
  emptyValueAllowed?: boolean;
  dataValue?: string;
}

type FloorSeedGroupedColumn = FloorSeedColumn & { columnGroupHeading: string };

interface FloorSeedSheet {
  title: string;
  liveTableName: string;
  letApiAccess: boolean;
  columns: readonly FloorSeedGroupedColumn[];
  endpoints?: Record<
    string,
    {
      heading: string;
      timeLastRan: FloorSeedColumn;
      runStatus: FloorSeedColumn;
    }
  >;
}

export const configSheetFloorSeed = {
  spreadsheetConfig: {
    title: "Spreadsheet Config",
    liveTableName: "spreadsheetConfig",
    letApiAccess: true,
    columns: [
      {
        header: "Table menu space",
        columnGroupHeading: "",
        columnType: "TEXT",
        dataValue: "Not used",
      },
    ],
    endpoints: {
      spreadsheetConfig_fillRowIdsTimeLastRan: {
        heading: "Fill Row IDs",
        timeLastRan: {
          header: "Fill row IDs, time last ran",
          columnType: "TEXT",
        },
        runStatus: {
          header: "Fill row IDs, run status",
          columnType: "TEXT",
        },
      },
      spreadsheetConfig_syncConfigSheetRowsTimeLastRan: {
        heading: "Sync Config Sheet Rows",
        timeLastRan: {
          header: "Sync config sheet rows, time last ran",
          columnType: "TEXT",
        },
        runStatus: {
          header: "Sync config sheet rows, run status",
          columnType: "TEXT",
        },
      },
    },
  },
  tableConfig: {
    title: "Table Config",
    liveTableName: "tableConfig",
    letApiAccess: true,
    columns: [
      {
        header: "Table ID",
        columnGroupHeading: "",
        columnType: "TEXT",
      },
      {
        header: "Table name",
        columnGroupHeading: "",
        columnType: "TEXT",
      },
      {
        header: "Sheet title",
        columnGroupHeading: "",
        columnType: "TEXT",
      },
      {
        header: "Let api access",
        columnGroupHeading: "",
        columnType: "BOOLEAN",
      },
    ],
  },
  columnConfig: {
    title: "Column Config",
    liveTableName: "columnConfig",
    letApiAccess: true,
    columns: [
      {
        header: "Table ID",
        columnGroupHeading: "",
        columnType: "TEXT",
      },
      {
        header: "Column ID",
        columnGroupHeading: "",
        columnType: "TEXT",
      },
      {
        header: "Table name",
        columnGroupHeading: "",
        columnType: "TEXT",
      },
      {
        header: "Header",
        columnGroupHeading: "",
        columnType: "TEXT",
      },
      {
        header: "Empty value allowed",
        columnGroupHeading: "",
        columnType: "BOOLEAN",
      },
    ],
  },
} as const satisfies Record<string, FloorSeedSheet>;

export type FloorTabName = keyof typeof configSheetFloorSeed;
type FloorTabSeed = (typeof configSheetFloorSeed)[FloorTabName];

export function isFloorTabName(name: string): name is FloorTabName {
  return Object.hasOwn(configSheetFloorSeed, name);
}

export function floorTabSeedByGid(sheetGid: number): FloorTabSeed | undefined {
  return floorTabSeedOf(tableConfigAloneOnGid(sheetGid));
}

export function floorTabSeedByTableId(
  tableId: string,
): FloorTabSeed | undefined {
  return floorTabSeedOf(tableConfigsByTableId().get(tableId));
}

function floorTabSeedOf(
  tableConfig: TableConfig | undefined,
): FloorTabSeed | undefined {
  const tableKey = floorTableKeyOf(tableConfig);
  if (tableKey === undefined) return undefined;
  return configSheetFloorSeed[tableKey];
}

function floorTableKeyOf(
  tableConfig: TableConfig | undefined,
): FloorTabName | undefined {
  if (tableConfig === undefined || !isFloorTabName(tableConfig.tableKey)) {
    return undefined;
  }
  return tableConfig.tableKey;
}

export function floorSeedColumns(
  tableName: FloorTabName,
): readonly FloorSeedColumn[] {
  if (tableName !== "spreadsheetConfig") {
    return configSheetFloorSeed[tableName].columns;
  }
  return [
    ...configSheetFloorSeed.spreadsheetConfig.columns,
    ...Obj.values(configSheetFloorSeed.spreadsheetConfig.endpoints).flatMap(
      (endpoint) => [endpoint.timeLastRan, endpoint.runStatus],
    ),
  ];
}

export function floorSeedColumnById(
  tableId: string,
  columnId: string,
): FloorSeedColumn | undefined {
  const tableKey = floorTableKeyOf(tableConfigsByTableId().get(tableId));
  if (tableKey === undefined) return undefined;
  return floorSeedColumnInSheet(tableKey, columnId);
}

export function floorSeedColumnInSheet(
  tableName: FloorTabName,
  columnId: string,
): FloorSeedColumn | undefined {
  return floorSeedColumns(tableName).find((seedColumn) => {
    const columnName = getSheetColumnNames(tableName).find(
      (name) =>
        getColumnTraitByName(tableName, name, "header") === seedColumn.header,
    );
    if (columnName === undefined) return false;
    return getColumnTraitByName(tableName, columnName, "columnId") === columnId;
  });
}

// A seed column lists emptyValueAllowed only when it's true.
export function floorSeedEmptyValueAllowed(
  seedColumn: FloorSeedColumn,
): boolean {
  return seedColumn.emptyValueAllowed ?? false;
}

export function floorColumnLabel(tableName: string, header: string): string {
  return `Floor column "${header}" on "${tableName}"`;
}
