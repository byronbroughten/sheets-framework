import type { CellValue } from "../../../src/00_Source/CellValues/cellValues";
import type { TableColumnType } from "../../../src/00_Source/RawSource/RawSource";
import { SheetIndex } from "../../../src/00_Source/RawSource/SheetIndex";
import { dimensionIds } from "../../../src/01_SpreadsheetSchema/dimensionIds";
import { tableLayout } from "../../../src/01_SpreadsheetSchema/tableLayout";
import { TableOrigin } from "../../../src/01_SpreadsheetSchema/TableOrigin";

export interface DevFixtureColumn {
  key: string;
  header: string;
  columnType: TableColumnType;
  emptyValueAllowed?: boolean;
  values: CellValue[];
  formula?: string;
}

export interface DevFixtureTable {
  tableName: string;
  idPrefix: string;
  origin: TableOrigin;
  entryCheckboxColumnKey?: string;
  columns: DevFixtureColumn[];
}

export interface DevFixtureSheet {
  sheetGid: number;
  title: string;
  tables: DevFixtureTable[];
}

// Pinned here, not read from sheets.config.json, so a bad config edit can't redirect the chore.
export const devSpreadsheetId = "19gIs4w8-2Nsin5zTN1TojOR1HiiT9Y-jCctC7doAMqM";

export function bodyRowCountOf(table: DevFixtureTable): number {
  return Math.max(...table.columns.map((column) => column.values.length));
}

function withIdColumn(table: DevFixtureTable): DevFixtureTable {
  const rowCount = bodyRowCountOf(table);
  const idColumn: DevFixtureColumn = {
    key: "id",
    header: tableLayout.idHeader,
    columnType: "TEXT",
    values: Array.from({ length: rowCount }, (_, index) =>
      dimensionIds.row(table.idPrefix, String(index + 1)),
    ),
  };
  return { ...table, columns: [idColumn, ...table.columns] };
}

function entryAmountColumns(
  entries: string[],
  amounts: number[],
): DevFixtureColumn[] {
  return [
    { key: "entry", header: "Entry", columnType: "TEXT", values: entries },
    { key: "amount", header: "Amount", columnType: "DOUBLE", values: amounts },
  ];
}

function nameColumn(values: string[]): DevFixtureColumn {
  return {
    key: "name",
    header: tableLayout.nameHeader,
    columnType: "TEXT",
    values,
  };
}

const expectedOrigin = TableOrigin.expected();

export const devFixtureSheets: readonly DevFixtureSheet[] = [
  {
    sheetGid: 1100001,
    title: "Item",
    tables: [
      withIdColumn({
        tableName: "item",
        idPrefix: "itm",
        origin: expectedOrigin,
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
    ],
  },
  {
    sheetGid: 1100002,
    title: "Value Types",
    tables: [
      withIdColumn({
        tableName: "valueTypes",
        idPrefix: "vty",
        origin: expectedOrigin,
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
    ],
  },
  {
    sheetGid: 1100003,
    title: "Log",
    tables: [
      {
        tableName: "log",
        idPrefix: "log",
        origin: expectedOrigin,
        columns: entryAmountColumns(["First entry", "Second entry"], [10, 20]),
      },
    ],
  },
  {
    sheetGid: 1100005,
    title: "Computed",
    tables: [
      {
        tableName: "computed",
        idPrefix: "cmp",
        origin: expectedOrigin,
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
    ],
  },
  {
    sheetGid: 1100004,
    title: "Run Item",
    tables: [
      withIdColumn({
        tableName: "runItem",
        idPrefix: "rit",
        origin: expectedOrigin,
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
    ],
  },
  {
    sheetGid: 1100006,
    title: "Dates",
    tables: [
      withIdColumn({
        tableName: "dates",
        idPrefix: "dat",
        origin: expectedOrigin,
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
    ],
  },
  // Right sits beside Left within Left's rows, and Below sits under Left within its columns.
  {
    sheetGid: 1100007,
    title: "Layout",
    tables: [
      {
        tableName: "layoutLeft",
        idPrefix: "lyl",
        origin: expectedOrigin,
        columns: entryAmountColumns(
          ["Left one", "Left two", "Left three"],
          [1, 2, 3],
        ),
      },
      {
        tableName: "layoutRight",
        idPrefix: "lyr",
        origin: new TableOrigin({
          headerRowIndex: SheetIndex.row(3),
          startColIndex: SheetIndex.col(3),
        }),
        columns: entryAmountColumns(["Right one", "Right two"], [10, 20]),
      },
    ],
  },
];
