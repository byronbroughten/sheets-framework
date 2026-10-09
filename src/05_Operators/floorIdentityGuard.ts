import { Obj } from "@byronbroughten/utils/obj";
import { Str } from "@byronbroughten/utils/str";

import {
  configSheetFloorSeed,
  floorColumnLabel,
  floorSeedColumns,
  type FloorTabName,
} from "../01_SpreadsheetSchema/configSheetFloorSeed";
import type {
  ColumnConfigsGeneric,
  TableConfigsBase,
} from "../01_SpreadsheetSchema/makeConfigs";

export interface FloorIdentitySource {
  tableConfigs: TableConfigsBase;
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
    ...floorTabIdentityChanges(previous.tableConfigs, next.tableConfigs),
    ...floorColumnIdentityChanges(previous.columnConfigs, next.columnConfigs),
  ];
  if (changes.length === 0) return;
  throw new Error(
    `The regen would change a floor identity, which every generated key built on it moves with. ${changes.join(" ")}`,
  );
}

function floorTabIdentityChanges(
  previous: TableConfigsBase,
  next: TableConfigsBase,
): string[] {
  return Obj.keys(configSheetFloorSeed).flatMap((tableKey) => {
    const previousConfig = previous[tableKey];
    const nextConfig = next[tableKey];
    if (previousConfig === undefined || nextConfig === undefined) return [];
    const changes: string[] = [];
    if (nextConfig.sheetGid !== previousConfig.sheetGid) {
      changes.push(
        `${floorTabLabel(tableKey)} GID was ${previousConfig.sheetGid} and is now ${nextConfig.sheetGid}.`,
      );
    }
    if (nextConfig.tableId !== previousConfig.tableId) {
      changes.push(
        `${floorTabLabel(tableKey)} Table ID was "${previousConfig.tableId}" and is now "${nextConfig.tableId}".`,
      );
    }
    if (nextConfig.idPrefix !== previousConfig.idPrefix) {
      changes.push(
        `${floorTabLabel(tableKey)} ID prefix was "${previousConfig.idPrefix}" and is now "${nextConfig.idPrefix}".`,
      );
    }
    return changes;
  });
}

function floorColumnIdentityChanges(
  previous: ColumnConfigsGeneric,
  next: ColumnConfigsGeneric,
): string[] {
  return Obj.keys(configSheetFloorSeed).flatMap((tableKey) => {
    return floorSeedColumns(tableKey).flatMap(({ header }) => {
      const columnName = Str.sentenceToCamelCase(header);
      const previousColumn = previous[tableKey]?.[columnName];
      const nextColumn = next[tableKey]?.[columnName];
      if (previousColumn === undefined || nextColumn === undefined) return [];
      if (nextColumn.columnId === previousColumn.columnId) return [];
      return [
        `${floorColumnLabel(tableKey, header)} had column ID "${previousColumn.columnId}" and is now "${nextColumn.columnId}".`,
      ];
    });
  });
}

function floorTabLabel(tableKey: FloorTabName): string {
  return `Floor tab "${tableKey}"`;
}
