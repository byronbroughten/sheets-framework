import { beforeEach, describe, expect, it } from "vitest";

import { installedConfigs } from "../01_SpreadsheetSchema/configRegister";
import { getSheetTraitByName } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  type FakeCell,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { SpreadsheetConfigOperator } from "./SpreadsheetConfigOperator";

const { columnConfigs, spreadsheetConfig } = installedConfigs();

const spreadsheetConfigGid = getSheetTraitByName(
  "spreadsheetConfig",
  "sheetGid",
);
const ssc = columnConfigs.spreadsheetConfig;
const spreadsheetConfigColumns = [
  ssc.idDelimiter,
  ssc.idHeader,
  ssc.nameHeader,
  ssc.startTableColumnIndexBase1,
  ssc.columnIdRowIndexBase1,
  ssc.columnGroupHeadingRowIndexBase1,
  ssc.actionRowIndexBase1,
  ssc.tableHeaderRowIndexBase1,
];
const columnIdRow = spreadsheetConfigColumns.map((column) => column.columnId);
const tableHeaderRow = spreadsheetConfigColumns.map((column) => column.header);

const columnIdRowIndex = spreadsheetConfig.columnIdRowIdxBase0;
const tableHeaderRowIndex = spreadsheetConfig.tableHeaderRowIndexBase0;
const firstDataRowIndex = tableHeaderRowIndex + 1;

const compiledValues = [":", "ID", "Name", 1, 1, 2, 3, 4] as const;

function refusal(...lines: string[]): string {
  return [
    'Spreadsheet Config layout values other than "ID header" and "Name header" are fixed; put back:',
    ...lines,
  ].join("\n");
}

function stubSpreadsheetConfigSheet(
  rowsByIndex: Record<number, FakeCell[]>,
  extraColumn?: { columnId: string; header: string },
) {
  const lastRowIndex = Math.max(
    tableHeaderRowIndex,
    ...Object.keys(rowsByIndex).map(Number),
  );
  return stubSheetsService({
    sheets: [
      {
        sheetId: spreadsheetConfigGid,
        title: "Spreadsheet Config",
        rows: buildGridRows({
          [columnIdRowIndex]: extraColumn
            ? [...columnIdRow, extraColumn.columnId]
            : columnIdRow,
          [tableHeaderRowIndex]: extraColumn
            ? [...tableHeaderRow, extraColumn.header]
            : tableHeaderRow,
          ...rowsByIndex,
        }),
        table: { endRowIndex: lastRowIndex + 1 },
      },
    ],
  });
}

function fetchedOperator() {
  const operator = SpreadsheetConfigOperator.init();
  operator.fetchLiveConfig();
  return operator;
}

beforeEach(() => {
  stubLogger();
});

describe("SpreadsheetConfigOperator.fetchLiveConfig / toFileSource", () => {
  it("emits makeSpreadsheetConfig of the first data row with base-1 indexes minus one", () => {
    stubSpreadsheetConfigSheet({ [firstDataRowIndex]: [...compiledValues] });

    expect(fetchedOperator().toFileSource("../makeConfigs")).toBe(
      [
        `import { makeSpreadsheetConfig } from "../makeConfigs";`,
        ``,
        `export const spreadsheetConfig = makeSpreadsheetConfig({`,
        `  idDelimiter: ":",`,
        `  idHeader: "ID",`,
        `  nameHeader: "Name",`,
        `  startTableColIndexBase0: 0,`,
        `  columnIdRowIdxBase0: 0,`,
        `  columnGroupHeadingRowIndexBase0: 1,`,
        `  actionRowIndexBase0: 2,`,
        `  tableHeaderRowIndexBase0: 3,`,
        `} as const);`,
        ``,
      ].join("\n"),
    );
  });

  it("emits an edited ID header as written", () => {
    stubSpreadsheetConfigSheet({
      [firstDataRowIndex]: [":", "Row ID", "Name", 1, 1, 2, 3, 4],
    });

    expect(fetchedOperator().toFileSource("../makeConfigs")).toContain(
      'idHeader: "Row ID"',
    );
  });

  it("emits an edited Name header as written", () => {
    stubSpreadsheetConfigSheet({
      [firstDataRowIndex]: [":", "ID", "Title", 1, 1, 2, 3, 4],
    });

    expect(fetchedOperator().toFileSource("../makeConfigs")).toContain(
      'nameHeader: "Title"',
    );
  });

  it("refuses an edited ID delimiter, naming it with the live and expected values", () => {
    stubSpreadsheetConfigSheet({
      [firstDataRowIndex]: ["|", "ID", "Name", 1, 1, 2, 3, 4],
    });

    expect(() => SpreadsheetConfigOperator.init().fetchLiveConfig()).toThrow(
      refusal('Spreadsheet Config column "ID delimiter" is "|"; expected ":".'),
    );
  });

  it("refuses every edited index in one error, in base 1", () => {
    stubSpreadsheetConfigSheet({
      [firstDataRowIndex]: [":", "ID", "Name", 2, 1, 2, 3, 5],
    });

    expect(() => SpreadsheetConfigOperator.init().fetchLiveConfig()).toThrow(
      refusal(
        'Spreadsheet Config column "Start table column index base 1" is 2; expected 1.',
        'Spreadsheet Config column "Table header row index base 1" is 5; expected 4.',
      ),
    );
  });

  it("ignores an extra column that is not in the closed map", () => {
    stubSpreadsheetConfigSheet(
      { [firstDataRowIndex]: [...compiledValues, "ignore me"] },
      { columnId: "notes", header: "Notes" },
    );

    expect(fetchedOperator().toFileSource("../makeConfigs")).toContain(
      'idDelimiter: ":"',
    );
    expect(fetchedOperator().toFileSource("../makeConfigs")).not.toContain(
      "Notes",
    );
  });

  it("throws when there are no data rows", () => {
    stubSpreadsheetConfigSheet({});

    expect(() => SpreadsheetConfigOperator.init().fetchLiveConfig()).toThrow(
      `Sheet "Spreadsheet Config" (gid ${spreadsheetConfigGid}) Table must have at least one data row.`,
    );
  });

  it("reads the first data row as the Table header row plus one, not a later filled row", () => {
    stubSpreadsheetConfigSheet({
      [firstDataRowIndex + 1]: [...compiledValues],
    });

    expect(() => SpreadsheetConfigOperator.init().fetchLiveConfig()).toThrow(
      `Column "idDelimiter" of sheet "spreadsheetConfig" is empty in row ${firstDataRowIndex}.`,
    );
  });

  it("does not treat a second filled row as a fetch-time error", () => {
    stubSpreadsheetConfigSheet({
      [firstDataRowIndex]: [...compiledValues],
      [firstDataRowIndex + 1]: ["x", "", "", "", "", "", ""],
    });

    expect(fetchedOperator().toFileSource("../makeConfigs")).toContain(
      'idDelimiter: ":"',
    );
  });

  it("throws when a guaranteed cell is blank", () => {
    stubSpreadsheetConfigSheet({
      [firstDataRowIndex]: ["", "ID", "Name", 1, 1, 2, 3, 4],
    });

    expect(() => SpreadsheetConfigOperator.init().fetchLiveConfig()).toThrow(
      `Column "idDelimiter" of sheet "spreadsheetConfig" is empty in row ${firstDataRowIndex}.`,
    );
  });

  it("refuses a non-numeric index, naming it", () => {
    stubSpreadsheetConfigSheet({
      [firstDataRowIndex]: [":", "ID", "Name", "one", 1, 2, 3, 4],
    });

    expect(() => SpreadsheetConfigOperator.init().fetchLiveConfig()).toThrow(
      refusal(
        'Spreadsheet Config column "Start table column index base 1" is "one"; expected 1.',
      ),
    );
  });
});
