import type { CellValue } from "../../../src/00_Source/CellValues/cellValues";
import type { TableColumnType } from "../../../src/00_Source/RawSource/RawSource";
import { dimensionIds } from "../../../src/01_SpreadsheetSchema/dimensionIds";
import { ssConfigGet } from "../../../src/01_SpreadsheetSchema/spreadsheetConfigTypes";

export interface DevFixtureColumn {
  key: string;
  header: string;
  columnType: TableColumnType;
  emptyValueAllowed?: boolean;
  values: CellValue[];
  formula?: string;
}

export interface DevFixtureSheet {
  sheetGid: number;
  title: string;
  tableName: string;
  idPrefix: string;
  entryCheckboxColumnKey?: string;
  columns: DevFixtureColumn[];
}

// Pinned here, not read from sheets.config.json, so a bad config edit can't redirect the chore.
export const devSpreadsheetId = "19gIs4w8-2Nsin5zTN1TojOR1HiiT9Y-jCctC7doAMqM";

function withIdColumn(sheet: DevFixtureSheet): DevFixtureSheet {
  const rowCount = Math.max(
    ...sheet.columns.map((column) => column.values.length),
  );
  const idColumn: DevFixtureColumn = {
    key: "id",
    header: ssConfigGet("idHeader"),
    columnType: "TEXT",
    values: Array.from({ length: rowCount }, (_, index) =>
      dimensionIds.row(sheet.idPrefix, String(index + 1)),
    ),
  };
  return { ...sheet, columns: [idColumn, ...sheet.columns] };
}

function nameColumn(values: string[]): DevFixtureColumn {
  return {
    key: "name",
    header: ssConfigGet("nameHeader"),
    columnType: "TEXT",
    values,
  };
}

// A function, not a const: the ID and Name headers come from the installed Spreadsheet Config.
export function devFixtureSheets(): DevFixtureSheet[] {
  return [
    withIdColumn({
      sheetGid: 1100001,
      title: "Item",
      tableName: "item",
      idPrefix: "itm",
      columns: [
        nameColumn(["Alpha", "Beta", "Gamma"]),
        {
          key: "optionalNote",
          header: "Optional note",
          columnType: "TEXT",
          emptyValueAllowed: true,
          values: ["", "A note", ""],
        },
        {
          key: "requiredCount",
          header: "Required count",
          columnType: "DOUBLE",
          emptyValueAllowed: false,
          values: [1, 2, 3],
        },
      ],
    }),
    withIdColumn({
      sheetGid: 1100002,
      title: "Value Types",
      tableName: "valueTypes",
      idPrefix: "vty",
      columns: [
        {
          key: "stringValue",
          header: "String value",
          columnType: "TEXT",
          values: ["One", "Two"],
        },
        {
          key: "numberValue",
          header: "Number value",
          columnType: "DOUBLE",
          values: [1.5, 2],
        },
        {
          key: "dateValue",
          header: "Date value",
          columnType: "DATE",
          values: [46000, 46001],
        },
        {
          key: "sampledBoolean",
          header: "Sampled boolean",
          columnType: "COLUMN_TYPE_UNSPECIFIED",
          values: [true, false],
        },
        {
          key: "checkbox",
          header: "Checkbox",
          columnType: "BOOLEAN",
          values: [true, false],
        },
      ],
    }),
    {
      sheetGid: 1100003,
      title: "Log",
      tableName: "log",
      idPrefix: "log",
      columns: [
        {
          key: "entry",
          header: "Entry",
          columnType: "TEXT",
          values: ["First entry", "Second entry"],
        },
        {
          key: "amount",
          header: "Amount",
          columnType: "DOUBLE",
          values: [10, 20],
        },
      ],
    },
    {
      sheetGid: 1100005,
      title: "Computed",
      tableName: "computed",
      idPrefix: "cmp",
      columns: [
        {
          key: "amount",
          header: "Amount",
          columnType: "DOUBLE",
          values: [1, 2],
        },
        {
          key: "rowNumber",
          header: "Row number",
          columnType: "DOUBLE",
          values: [],
          formula: "=ROW()",
        },
      ],
    },
    withIdColumn({
      sheetGid: 1100004,
      title: "Run Item",
      tableName: "runItem",
      idPrefix: "rit",
      entryCheckboxColumnKey: "result",
      columns: [
        nameColumn(["First", "Second", "Third"]),
        {
          key: "selected",
          header: "Selected",
          columnType: "BOOLEAN",
          values: [false, false, false],
        },
        {
          key: "result",
          header: "Result",
          columnType: "TEXT",
          values: ["", "", ""],
        },
        {
          key: "startTime",
          header: "Start time",
          columnType: "TEXT",
          values: ["", "", ""],
        },
        {
          key: "runStatus",
          header: "Run status",
          columnType: "TEXT",
          values: ["", "", ""],
        },
      ],
    }),
    withIdColumn({
      sheetGid: 1100006,
      title: "Dates",
      tableName: "dates",
      idPrefix: "dat",
      columns: [
        {
          key: "requiredDate",
          header: "Required date",
          columnType: "DATE",
          emptyValueAllowed: false,
          values: [46000, 46001],
        },
        {
          key: "optionalDate",
          header: "Optional date",
          columnType: "DATE",
          emptyValueAllowed: true,
          values: ["", 46002],
        },
      ],
    }),
  ];
}
