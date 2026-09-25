import {
  configSheetFloorSeed,
  floorColumnLabel,
  floorSeedColumns,
  type FloorTabName,
} from "../01_SpreadsheetSchema/configSheetFloorSeed";
import type {
  ColumnConfigsGeneric,
  SheetConfigsBase,
} from "../01_SpreadsheetSchema/makeConfigs";
import { Obj } from "../utils/Obj";
import { Str } from "../utils/Str";

export interface FloorIdentitySource {
  sheetConfigs: SheetConfigsBase;
  columnConfigs: ColumnConfigsGeneric;
}

export function assertFloorIdentityUnchanged({
  previous,
  next,
}: {
  previous: FloorIdentitySource;
  next: FloorIdentitySource;
}): void {
  const changes = [
    ...floorSheetIdentityChanges(previous.sheetConfigs, next.sheetConfigs),
    ...floorColumnIdentityChanges(previous.columnConfigs, next.columnConfigs),
  ];
  if (changes.length === 0) return;
  throw new Error(
    `The regen would change a floor identity, which every generated key built on it moves with. ${changes.join(" ")}`,
  );
}

function floorSheetIdentityChanges(
  previous: SheetConfigsBase,
  next: SheetConfigsBase,
): string[] {
  return Obj.keys(configSheetFloorSeed).flatMap((sheetName) => {
    const previousConfig = previous[sheetName];
    const nextConfig = next[sheetName];
    if (previousConfig === undefined || nextConfig === undefined) return [];
    const changes: string[] = [];
    if (nextConfig.sheetGid !== previousConfig.sheetGid) {
      changes.push(
        `${floorTabLabel(sheetName)} GID was ${previousConfig.sheetGid} and is now ${nextConfig.sheetGid}.`,
      );
    }
    if (nextConfig.idPrefix !== previousConfig.idPrefix) {
      changes.push(
        `${floorTabLabel(sheetName)} ID prefix was "${previousConfig.idPrefix}" and is now "${nextConfig.idPrefix}".`,
      );
    }
    return changes;
  });
}

function floorColumnIdentityChanges(
  previous: ColumnConfigsGeneric,
  next: ColumnConfigsGeneric,
): string[] {
  return Obj.keys(configSheetFloorSeed).flatMap((sheetName) => {
    return floorSeedColumns(sheetName).flatMap(({ header }) => {
      const columnName = Str.sentenceToCamelCase(header);
      const previousColumn = previous[sheetName]?.[columnName];
      const nextColumn = next[sheetName]?.[columnName];
      if (previousColumn === undefined || nextColumn === undefined) return [];
      if (nextColumn.columnId === previousColumn.columnId) return [];
      return [
        `${floorColumnLabel(sheetName, header)} had column ID "${previousColumn.columnId}" and is now "${nextColumn.columnId}".`,
      ];
    });
  });
}

function floorTabLabel(sheetName: FloorTabName): string {
  return `Floor tab "${sheetName}"`;
}
