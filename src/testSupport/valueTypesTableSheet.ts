import {
  type Configs,
  installedConfigs,
} from "../01_SpreadsheetSchema/configReaders/configRegister";
import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import {
  type FakeSheetProperties,
  type FakeSheetsService,
  stubSheetsService,
} from "./fakeSheetsService";
import {
  type FakeBodyRow,
  fakeTableSheet,
} from "./fakeSheetsService/fakeTableSheet";

type ValueTypesColumnConfigs = Configs["columnConfigs"]["valueTypes"];

export const valueTypesGid = getTableTraitByName("valueTypes", "sheetGid");

export type ValueTypesColumnName = keyof ValueTypesColumnConfigs & string;

export type ValueTypesBodyRow = FakeBodyRow<ValueTypesColumnName>;

function valueTypesColumnConfigs(): ValueTypesColumnConfigs {
  return installedConfigs().columnConfigs.valueTypes;
}

export function valueTypesColumnNames(): ValueTypesColumnName[] {
  return Object.keys(valueTypesColumnConfigs()) as ValueTypesColumnName[];
}

/** Values in the id, text and checkbox columns. */
export const filledValueTypesRow: ValueTypesBodyRow = {
  id: "r:vty:filled",
  stringValue: "Item",
  checkbox: true,
};

/** The blank row: nothing in any column. */
export const blankValueTypesRow: ValueTypesBodyRow = {};

export interface ValueTypesTableSheetOptions {
  columnNames?: readonly ValueTypesColumnName[]; // Only these are on the sheet; every configured one by default.
  rowsWithNoGridData?: number[];
}

/** One sheet holding the Value Types Table, its body rows from the top data row down. */
export function valueTypesTableSheet(
  bodyRows: readonly ValueTypesBodyRow[],
  options: ValueTypesTableSheetOptions = {},
): FakeSheetProperties {
  return {
    ...fakeTableSheet.build({
      sheetId: valueTypesGid,
      title: "Value Types",
      columnConfigs: valueTypesColumnConfigs(),
      columnNames: options.columnNames ?? valueTypesColumnNames(),
      bodyRows,
    }),
    rowsWithNoGridData: options.rowsWithNoGridData,
  };
}

/** Stubs the Sheets service with just the Value Types Table's sheet. */
export function stubValueTypesTable(
  bodyRows: readonly ValueTypesBodyRow[],
  options?: ValueTypesTableSheetOptions,
): FakeSheetsService {
  return stubSheetsService({
    sheets: [valueTypesTableSheet(bodyRows, options)],
  });
}
