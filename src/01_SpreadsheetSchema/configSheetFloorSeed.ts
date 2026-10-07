import type { TableColumnType } from "../00_Source/RawSource/RawSource";
import { Obj } from "../utils/Obj";
import {
  getColumnTraitByName,
  getSheetColumnNames,
} from "./columnConfigsTypes";
import {
  type TableConfig,
  tableConfigsByGid,
  tableConfigsByTableId,
} from "./tableConfigsTypes";

export type FloorColumnType = Extract<
  TableColumnType,
  "TEXT" | "DOUBLE" | "BOOLEAN"
>;

export interface FloorSeedColumn {
  header: string;
  columnType: FloorColumnType;
  emptyValueAllowed: boolean;
  dataValue?: string;
}

type FloorSeedGroupedColumn = FloorSeedColumn & { columnGroupHeading: string };

interface FloorSeedSheet {
  title: string;
  tableName: string;
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
  exampleColumn?: {
    header: string;
    columnType: FloorColumnType;
    seededValues: readonly string[];
  };
}

export const configSheetFloorSeed = {
  spreadsheetConfig: {
    title: "Spreadsheet Config",
    tableName: "spreadsheetConfig",
    letApiAccess: true,
    columns: [
      {
        header: "Table menu space",
        columnGroupHeading: "",
        columnType: "TEXT",
        emptyValueAllowed: false,
        dataValue: "Not used",
      },
    ],
    endpoints: {
      spreadsheetConfig_fillRowIdsTimeLastRan: {
        heading: "Fill Row IDs",
        timeLastRan: {
          header: "Fill row IDs, time last ran",
          columnType: "TEXT",
          emptyValueAllowed: false,
        },
        runStatus: {
          header: "Fill row IDs, run status",
          columnType: "TEXT",
          emptyValueAllowed: false,
        },
      },
      spreadsheetConfig_syncConfigSheetRowsTimeLastRan: {
        heading: "Sync Config Sheet Rows",
        timeLastRan: {
          header: "Sync config sheet rows, time last ran",
          columnType: "TEXT",
          emptyValueAllowed: false,
        },
        runStatus: {
          header: "Sync config sheet rows, run status",
          columnType: "TEXT",
          emptyValueAllowed: false,
        },
      },
    },
  },
  tableConfig: {
    title: "Table Config",
    tableName: "tableConfig",
    letApiAccess: true,
    columns: [
      {
        header: "Table ID",
        columnGroupHeading: "",
        columnType: "TEXT",
        emptyValueAllowed: false,
      },
      {
        header: "Table name",
        columnGroupHeading: "",
        columnType: "TEXT",
        emptyValueAllowed: false,
      },
      {
        header: "Sheet title",
        columnGroupHeading: "",
        columnType: "TEXT",
        emptyValueAllowed: false,
      },
      {
        header: "Let api access",
        columnGroupHeading: "",
        columnType: "BOOLEAN",
        emptyValueAllowed: false,
      },
    ],
  },
  columnConfig: {
    title: "Column Config",
    tableName: "columnConfig",
    letApiAccess: true,
    columns: [
      {
        header: "Table ID",
        columnGroupHeading: "",
        columnType: "TEXT",
        emptyValueAllowed: false,
      },
      {
        header: "Column ID",
        columnGroupHeading: "",
        columnType: "TEXT",
        emptyValueAllowed: false,
      },
      {
        header: "Table name",
        columnGroupHeading: "",
        columnType: "TEXT",
        emptyValueAllowed: false,
      },
      {
        header: "Header",
        columnGroupHeading: "",
        columnType: "TEXT",
        emptyValueAllowed: false,
      },
      {
        header: "Empty value allowed",
        columnGroupHeading: "",
        columnType: "BOOLEAN",
        emptyValueAllowed: false,
      },
    ],
  },
  valueConfig: {
    title: "Value Config",
    tableName: "valueConfig",
    letApiAccess: true,
    columns: [],
    exampleColumn: {
      header: "Example value",
      columnType: "TEXT",
      seededValues: ["Example one", "Example two"],
    },
  },
} as const satisfies Record<string, FloorSeedSheet>;

export type FloorTabName = keyof typeof configSheetFloorSeed;
type FloorTabSeed = (typeof configSheetFloorSeed)[FloorTabName];

export function isFloorTabName(name: string): name is FloorTabName {
  return Object.hasOwn(configSheetFloorSeed, name);
}

export function floorTabSeedByGid(sheetGid: number): FloorTabSeed | undefined {
  return floorTabSeedOf(tableConfigsByGid().get(sheetGid));
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
  sheetName: FloorTabName,
): readonly FloorSeedColumn[] {
  if (sheetName === "valueConfig") return [];
  if (sheetName !== "spreadsheetConfig") {
    return configSheetFloorSeed[sheetName].columns;
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
  sheetName: FloorTabName,
  columnId: string,
): FloorSeedColumn | undefined {
  return floorSeedColumns(sheetName).find((seedColumn) => {
    const columnName = getSheetColumnNames(sheetName).find(
      (name) =>
        getColumnTraitByName(sheetName, name, "header") === seedColumn.header,
    );
    if (columnName === undefined) return false;
    return getColumnTraitByName(sheetName, columnName, "columnId") === columnId;
  });
}

export function floorColumnLabel(sheetName: string, header: string): string {
  return `Floor column "${header}" on "${sheetName}"`;
}
