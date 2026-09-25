import {
  type ColumnName,
  getColumnTraitByName,
  getSheetColumnNames,
} from "../../01_SpreadsheetSchema/columnConfigsTypes";
import {
  configSheetFloorSeed,
  type FloorTabName,
} from "../../01_SpreadsheetSchema/configSheetFloorSeed";
import { Obj } from "../../utils/Obj";

export type FloorSheetName = Exclude<FloorTabName, "valueConfig">;

export interface FloorColumnRestore {
  header: string;
  columnId: string;
  groupHeading: string;
}

export function columnNameByHeader<SN extends FloorSheetName>(
  sheetName: SN,
  header: string,
): ColumnName<SN> {
  const columnName = getSheetColumnNames(sheetName).find(
    (name) => getColumnTraitByName(sheetName, name, "header") === header,
  );
  if (columnName === undefined) {
    throw new Error(
      `Floor seed header ${JSON.stringify(header)} is not a column on ${sheetName}.`,
    );
  }
  return columnName;
}

export function floorSheetNames(): FloorSheetName[] {
  return Obj.keys(configSheetFloorSeed).filter(
    (sheetName): sheetName is FloorSheetName => sheetName !== "valueConfig",
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

export function floorColumnsToRestore<SN extends FloorSheetName>(
  sheetName: SN,
): FloorColumnRestore[] {
  const seedColumns = configSheetFloorSeed[sheetName].columns.map((column) =>
    floorColumnRestore(sheetName, {
      header: column.header,
      groupHeading: column.columnGroupHeading,
    }),
  );
  if (sheetName !== "spreadsheetConfig") return seedColumns;
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

export function floorColumnRestore<SN extends FloorSheetName>(
  sheetName: SN,
  { header, groupHeading }: Pick<FloorColumnRestore, "header" | "groupHeading">,
): FloorColumnRestore {
  const columnName = columnNameByHeader(sheetName, header);
  return {
    header,
    columnId: getColumnTraitByName(sheetName, columnName, "columnId"),
    groupHeading,
  };
}
