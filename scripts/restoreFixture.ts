// The scripts project installs no configs, and valueTypesTableSheet reads them as it loads; each test imports them first too.
import "../dev/installDevConfigs.ts";

import { tableLayout } from "../src/01_SpreadsheetSchema/tableLayout.ts";
import type { FakeSheetsService } from "../src/testSupport/fakeSheetsService.ts";
import {
  type ValueTypesColumnName,
  valueTypesColumnNames,
} from "../src/testSupport/valueTypesTableSheet.ts";
import type { ConfirmedSend } from "./choreFixture.ts";

export const topDataRowIndex = tableLayout.headerZoneDepth;

// The Value Types Table's columns, with its ID column or without.
export interface TableCase {
  columnNames: ValueTypesColumnName[];
}

export const withRowIds: TableCase = { columnNames: valueTypesColumnNames() };
export const withoutRowIds: TableCase = {
  columnNames: valueTypesColumnNames().filter((name) => name !== "id"),
};

export function colOf(
  { columnNames }: TableCase,
  name: ValueTypesColumnName,
): number {
  return columnNames.indexOf(name);
}

// A confirmed send over the fake, and the restore its undo line names.
export interface SentCase {
  fake: FakeSheetsService;
  restore: ConfirmedSend["restore"];
}
