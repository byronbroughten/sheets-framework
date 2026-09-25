import { Obj } from "../utils/Obj";
import {
  configSheetFloorSeed,
  floorColumnLabel,
  type FloorColumnType,
  type FloorSeedColumn,
  floorSeedColumnInSheet,
  floorSeedColumns,
} from "./configSheetFloorSeed";
import type {
  ColumnConfigsGeneric,
  ColumnConfigStored,
  SheetConfigsBase,
} from "./makeConfigs";
import type { ValueName } from "./valueSchemas";

const floorTabNames = Obj.keys(configSheetFloorSeed);

const floorColumnTypeValueNames: Record<FloorColumnType, ValueName> = {
  TEXT: "string",
  DOUBLE: "number",
  BOOLEAN: "checkbox",
};

// Resolves column IDs through the last generated configs, so it's sound only after the floor identity guard.
export function assertFloorMatchesSeed(
  sheetConfigs: SheetConfigsBase,
  columnConfigs: ColumnConfigsGeneric,
): void {
  assertFloorTabEntries(sheetConfigs);
  assertFloorColumnEntries(columnConfigs);
}

function assertFloorTabEntries(sheetConfigs: SheetConfigsBase): void {
  const missing = floorTabNames.find(
    (sheetName) => !Object.hasOwn(sheetConfigs, sheetName),
  );
  if (missing === undefined) return;
  throw new Error(
    `Floor tab "${missing}" has no floor entry; Let api access may be unticked on its Sheet Config row.`,
  );
}

function assertFloorColumnEntries(columnConfigs: ColumnConfigsGeneric): void {
  const mismatches = floorTabNames.flatMap((sheetName) => {
    const matched = new Set<FloorSeedColumn>();
    const entryMismatches = Object.entries(
      columnConfigs[sheetName] ?? {},
    ).flatMap(([, column]) => {
      const seedColumn = floorSeedColumnInSheet(sheetName, column.columnId);
      if (seedColumn === undefined) return [];
      matched.add(seedColumn);
      const label = `${floorColumnLabel(sheetName, seedColumn.header)} (column ID "${column.columnId}")`;
      return floorColumnMismatches(label, column, seedColumn);
    });
    const missing = floorSeedColumns(sheetName)
      .filter((seedColumn) => !matched.has(seedColumn))
      .map(
        ({ header }) =>
          `${floorColumnLabel(sheetName, header)} has no floor entry.`,
      );
    return [...entryMismatches, ...missing];
  });
  if (mismatches.length === 0) return;
  throw new Error(
    `The generated configs differ from the floor seed. ${mismatches.join(" ")}`,
  );
}

function floorColumnMismatches(
  label: string,
  column: ColumnConfigStored,
  seedColumn: FloorSeedColumn,
): string[] {
  const mismatches: string[] = [];
  if (column.header !== seedColumn.header) {
    mismatches.push(
      `${label} has header "${column.header}" where the floor seed has "${seedColumn.header}".`,
    );
  }
  const seedValueName = floorColumnTypeValueNames[seedColumn.columnType];
  if (column.valueName !== seedValueName) {
    mismatches.push(
      `${label} has valueName "${column.valueName}" where the floor seed's column type ${seedColumn.columnType} implies "${seedValueName}".`,
    );
  }
  if (column.emptyValueAllowed !== seedColumn.emptyValueAllowed) {
    mismatches.push(
      `${label} has emptyValueAllowed ${column.emptyValueAllowed} where the floor seed has ${seedColumn.emptyValueAllowed}.`,
    );
  }
  return mismatches;
}
