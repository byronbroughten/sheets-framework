import { Obj } from "@byronbroughten/utils/obj";

import {
  type ColumnName,
  getColumnTraitByName,
  getSheetColumnNames,
} from "../../01_SpreadsheetSchema/columnConfigsTypes";
import {
  configSheetFloorSeed,
  type FloorTabName,
} from "../../01_SpreadsheetSchema/configSheetFloorSeed";

export type FloorSheetName = Exclude<FloorTabName, "valueConfig">;

export interface FloorColumnRestore {
  header: string;
  columnId: string;
  groupHeading: string;
}

export function columnNameByHeader<TN extends FloorSheetName>(
  tableName: TN,
  header: string,
): ColumnName<TN> {
  const columnName = getSheetColumnNames(tableName).find(
    (name) => getColumnTraitByName(tableName, name, "header") === header,
  );
  if (columnName === undefined) {
    throw new Error(
      `Floor seed header ${JSON.stringify(header)} is not a column on ${tableName}.`,
    );
  }
  return columnName;
}

export function floorSheetNames(): FloorSheetName[] {
  return Obj.keys(configSheetFloorSeed).filter(
    (tableName): tableName is FloorSheetName => tableName !== "valueConfig",
  );
}

export function spreadsheetConfigFeedbackColumnNames(): ColumnName<"spreadsheetConfig">[] {
  return Obj.values(configSheetFloorSeed.spreadsheetConfig.endpoints).flatMap(
    (endpoint) => [
      columnNameByHeader("spreadsheetConfig", endpoint.timeLastRan.header),
      columnNameByHeader("spreadsheetConfig", endpoint.runStatus.header),
    ],
  );
}

export function floorColumnsToRestore<TN extends FloorSheetName>(
  tableName: TN,
): FloorColumnRestore[] {
  const seedColumns = configSheetFloorSeed[tableName].columns.map((column) =>
    floorColumnRestore(tableName, {
      header: column.header,
      groupHeading: column.columnGroupHeading,
    }),
  );
  if (tableName !== "spreadsheetConfig") return seedColumns;
  const endpointColumns = Obj.values(
    configSheetFloorSeed.spreadsheetConfig.endpoints,
  ).flatMap((endpoint) =>
    [endpoint.timeLastRan, endpoint.runStatus].map((column) =>
      floorColumnRestore("spreadsheetConfig", {
        header: column.header,
        groupHeading: endpoint.heading,
      }),
    ),
  );
  return [...seedColumns, ...endpointColumns];
}

export function floorColumnRestore<TN extends FloorSheetName>(
  tableName: TN,
  { header, groupHeading }: Pick<FloorColumnRestore, "header" | "groupHeading">,
): FloorColumnRestore {
  const columnName = columnNameByHeader(tableName, header);
  return {
    header,
    columnId: getColumnTraitByName(tableName, columnName, "columnId"),
    groupHeading,
  };
}
