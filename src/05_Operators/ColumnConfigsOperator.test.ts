import type { StrictOmit } from "@byronbroughten/utils/obj";
import { beforeEach, describe, expect, it } from "vitest";

import { installedConfigs } from "../01_SpreadsheetSchema/configReaders/configRegister";
import type { ColumnConfigStored } from "../01_SpreadsheetSchema/configReaders/makeConfigs";
import type { TableReading } from "../01_SpreadsheetSchema/configReaders/SyncPlanner";
import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  type FakeCell,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { ColumnConfigsOperator } from "./ColumnConfigsOperator";

const { columnConfigs } = installedConfigs();

// The previous configs grant API access, so the fixtures reuse Table IDs the installed configs hold.
const itemTableId = getTableTraitByName("item", "tableId");
const logTableId = getTableTraitByName("log", "tableId");
const itemSheetGid = getTableTraitByName("item", "sheetGid");
const logSheetGid = getTableTraitByName("log", "sheetGid");
const ungrantedGid = 999009;

beforeEach(() => {
  stubLogger();
});

function syncedOperator(): ColumnConfigsOperator {
  const operator = ColumnConfigsOperator.init();
  operator.ss.fetchAllSheetProperties();
  operator.fetchHeadRows();
  operator.addMissingColumnIds().recordUntypedColumns();
  return operator;
}

function stubWidgetAndBrandNewSheet(): void {
  stubSheetsService({
    sheets: [
      {
        sheetId: itemSheetGid,
        title: "Widget",
        rows: buildGridRows({
          0: ["c:itm:aaa", "c:itm:bbb"],
          3: ["Unit Price", "Notes"],
          4: [42, "a note"],
        }),
        table: { tableId: itemTableId, name: "Widget", endRowIndex: 5 },
      },
      {
        sheetId: logSheetGid,
        title: "Brand New Sheet",
        rows: buildGridRows({
          0: ["c:log:ccc"],
          3: ["Some Field"],
          4: ["x"],
        }),
        table: { tableId: logTableId, name: "Brand New Sheet", endRowIndex: 5 },
      },
    ],
  });
}

describe("ColumnConfigsOperator.newColumnConfigs / toFileSource", () => {
  const emptyBaseline = {
    formula: null,
    columnType: { type: null, rule: null },
    actionRowCell: null,
    groupHeading1: null,
  };

  it("groups columns by Table key", () => {
    stubWidgetAndBrandNewSheet();

    const entries = syncedOperator().newColumnConfigs();

    expect(entries.widget).toEqual({
      unitPrice: {
        columnId: "c:itm:aaa",
        valueName: "number",
        header: "Unit Price",
        isFormula: false,
        emptyValueAllowed: false,
        customDefaultValue: null,
        ...emptyBaseline,
      },
      notes: {
        columnId: "c:itm:bbb",
        valueName: "string",
        header: "Notes",
        isFormula: false,
        emptyValueAllowed: false,
        customDefaultValue: null,
        ...emptyBaseline,
      },
    });
    expect(entries.brandNewSheet).toEqual({
      someField: {
        columnId: "c:log:ccc",
        valueName: "string",
        header: "Some Field",
        isFormula: false,
        emptyValueAllowed: false,
        customDefaultValue: null,
        ...emptyBaseline,
      },
    });
  });

  it("emits one labeled column-config record per line", () => {
    stubWidgetAndBrandNewSheet();

    const sourceLines = syncedOperator().toFileSource().split("\n");

    expect(sourceLines).toContain(
      '    "unitPrice": { "columnId": "c:itm:aaa", "header": "Unit Price", "valueName": "number", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":null,"rule":null}, "actionRowCell": null, "groupHeading1": null },',
    );
    expect(
      sourceLines.find((line) => line.includes('"widget":')),
    ).not.toContain('"columnId"');
    for (const line of sourceLines.filter((line) =>
      line.includes('"columnId"'),
    )) {
      expect(line).toContain('"header"');
      expect(line).toContain('"valueName"');
      expect(line).toContain('"isFormula"');
      expect(line).toContain('"emptyValueAllowed"');
      expect(line).toContain('"customDefaultValue"');
    }
  });

  describe("empty value allowed", () => {
    type EmptyValueAllowedTrait = Pick<
      ColumnConfigStored,
      "columnId" | "emptyValueAllowed"
    >;
    function traitOf({
      columnId,
      emptyValueAllowed,
    }: EmptyValueAllowedTrait): EmptyValueAllowedTrait {
      return { columnId, emptyValueAllowed };
    }
    const previousColumns = Object.values(columnConfigs.item);
    const previousTraits = previousColumns.map(traitOf);

    function stubItemWithNewColumn(): void {
      stubSheetsService({
        sheets: [
          {
            sheetId: itemSheetGid,
            title: "Item",
            rows: buildGridRows({
              0: [
                ...previousColumns.map(({ columnId }) => columnId),
                "c:itm:new",
              ],
              3: [...previousColumns.map(({ header }) => header), "Brand New"],
            }),
            table: { tableId: itemTableId, name: "Item", endRowIndex: 5 },
          },
        ],
      });
    }

    function emittedTraits(
      operator: ColumnConfigsOperator,
    ): EmptyValueAllowedTrait[] {
      return Object.values(operator.newColumnConfigs().item ?? {}).map(traitOf);
    }

    it("carries each column's trait forward from the previous configs, and requires a new column", () => {
      stubItemWithNewColumn();

      expect(emittedTraits(syncedOperator())).toEqual([
        ...previousTraits,
        { columnId: "c:itm:new", emptyValueAllowed: false },
      ]);
    });

    it("takes a recorded trait first, falling back to the previous configs for the rest", () => {
      stubItemWithNewColumn();
      const operator = syncedOperator();

      operator.recordEmptyValueAllowed({
        tables: {
          item: {
            traits: {},
            columns: { brandNew: { emptyValueAllowed: true } },
          },
        },
        tabs: {},
      });

      expect(emittedTraits(operator)).toEqual([
        ...previousTraits,
        { columnId: "c:itm:new", emptyValueAllowed: true },
      ]);
    });
  });

  it("throws when two headers in the same Table camelCase to the same column name", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: itemSheetGid,
          title: "Widget",
          rows: buildGridRows({
            0: ["c:itm:aaa", "c:itm:bbb"],
            3: ["Unit Price", "Unit  Price"],
            4: [1, 2],
          }),
          table: { tableId: itemTableId, name: "Widget", endRowIndex: 5 },
        },
      ],
    });

    expect(() => syncedOperator().newColumnConfigs()).toThrow(
      /duplicate column name "unitPrice"/,
    );
  });
});

describe("ColumnConfigsOperator.traitReadings", () => {
  it("reads emptyValueAllowed from the previous configs on both sides and header and name from each side, by previous names", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: itemSheetGid,
          title: "Widget",
          rows: buildGridRows({
            0: Object.values(columnConfigs.item).map(
              ({ columnId }) => columnId,
            ),
            3: Object.values(columnConfigs.item).map(
              ({ header }) => `${header} Renamed`,
            ),
          }),
          table: { tableId: itemTableId, name: "Widget", endRowIndex: 5 },
        },
      ],
    });
    function expectedColumns(headerSuffix: string): TableReading["columns"] {
      return Object.fromEntries(
        Object.entries(columnConfigs.item).map(
          ([columnName, { header, emptyValueAllowed }]) => [
            columnName,
            { header: `${header}${headerSuffix}`, emptyValueAllowed },
          ],
        ),
      );
    }

    const { previous, live } = syncedOperator().traitReadings();

    expect(previous.tables.item).toEqual({
      traits: { name: getTableTraitByName("item", "tableName") },
      columns: expectedColumns(""),
    });
    expect(live.tables.item).toEqual({
      traits: { name: "Widget" },
      columns: expectedColumns(" Renamed"),
    });
    expect(Object.keys(live.tables)).toEqual(["item"]);
  });

  it("names a column new to the configs by its header, and requires it", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: itemSheetGid,
          title: "Item",
          rows: buildGridRows({ 0: ["c:itm:new"], 3: ["Brand New"] }),
          table: { tableId: itemTableId, name: "Item", endRowIndex: 5 },
        },
      ],
    });

    const { live } = syncedOperator().traitReadings();

    expect(live.tables.item?.columns).toEqual({
      brandNew: { header: "Brand New", emptyValueAllowed: false },
    });
  });
});

describe("ColumnConfigsOperator.newColumnConfigs value-source and live-fact detection", () => {
  it("detects a value source from the column's live data-validation formula", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: itemSheetGid,
          title: "Item",
          rows: buildGridRows({
            0: ["c:itm:corr03", "c:itm:corr05"],
            3: ["Description", "Kind"],
            4: ["Unit (base)", "Unit (base)"],
          }),
          table: {
            tableId: itemTableId,
            name: "Item",
            endRowIndex: 5,
            columnValidationValues: { 0: ["=Item[Kind]"] },
          },
        },
      ],
    });

    expect(
      syncedOperator().newColumnConfigs().item?.description?.valueName,
    ).toBe("itemKind");
  });

  it("detects a live formula and a date-formatted number", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: itemSheetGid,
          title: "Item",
          rows: buildGridRows({
            0: ["c:itm:corr04"],
            3: ["Due-by Date"],
            4: [{ value: 45000, isFormula: true, numberFormatType: "DATE" }],
          }),
          table: { tableId: itemTableId, name: "Item", endRowIndex: 5 },
        },
      ],
    });

    expect(syncedOperator().newColumnConfigs().item?.dueByDate).toMatchObject({
      valueName: "date",
      isFormula: true,
    });
  });
});

interface ColumnsUnderTest {
  headers: string[];
  topDataRow?: FakeCell[];
  topDataRowAbsence?: "rowsWithNoGridData" | "rowsWithNoGridBlock";
  columnTypes?: Record<number, string>;
  columnValidationValues?: Record<number, string[]>;
  columnValidationConditionTypes?: Record<number, string>;
}

function syncColumnsUnderTest({
  headers,
  topDataRow = [],
  topDataRowAbsence,
  columnTypes,
  columnValidationValues,
  columnValidationConditionTypes,
}: ColumnsUnderTest): ColumnConfigsOperator {
  const columnIds = headers.map((_, index) => `c:itm:col${index}`);
  stubSheetsService({
    sheets: [
      {
        sheetId: itemSheetGid,
        title: "Item",
        rows: buildGridRows({ 0: columnIds, 3: headers, 4: topDataRow }),
        ...(topDataRowAbsence ? { [topDataRowAbsence]: [4] } : {}),
        table: {
          tableId: itemTableId,
          name: "Item",
          endRowIndex: 5,
          columnTypes,
          columnValidationValues,
          columnValidationConditionTypes,
        },
      },
    ],
  });
  return syncedOperator();
}

function valueTitles(operator: ColumnConfigsOperator, count: number) {
  const { profile } = operator.ss.raw.table(itemTableId);
  const titles = profile.columnIds.map((columnId) =>
    profile.columnById(columnId).valueTitle(),
  );
  expect(titles).toHaveLength(count);
  return titles;
}

// The reported bug's shape: a sheet emptied by a run that consumed its input.
function syncBlankSheetUnderTest(
  props: StrictOmit<ColumnsUnderTest, "topDataRow">,
): ColumnConfigsOperator {
  return syncColumnsUnderTest({
    topDataRowAbsence: "rowsWithNoGridData",
    ...props,
  });
}

describe("ColumnConfigsOperator.recordUntypedColumns -> declared column types", () => {
  it("maps every declared column type to its value name", () => {
    const operator = syncColumnsUnderTest({
      headers: [
        "Amount",
        "Count",
        "Rate",
        "Moved In",
        "Start Time",
        "Updated At",
        "Notes",
        "Owner",
        "Active",
      ],
      columnTypes: {
        0: "CURRENCY",
        1: "DOUBLE",
        2: "PERCENT",
        3: "DATE",
        4: "TIME",
        5: "DATE_TIME",
        6: "TEXT",
        7: "PEOPLE_CHIP",
        8: "BOOLEAN",
      },
    });

    expect(valueTitles(operator, 9)).toEqual([
      "number",
      "number",
      "number",
      "date",
      "number",
      "number",
      "string",
      "string",
      "checkbox",
    ]);
  });

  it("counts an undeclared column holding a boolean as a guessed boolean, not a checkbox", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Active"],
      topDataRow: [true],
    });

    expect(valueTitles(operator, 1)).toEqual(["boolean"]);
    expect(operator.untypedColumnsSummary()).toContain(
      "1 column(s) across 1 Table(s)",
    );
  });

  it("treats BOOLEAN Table validation with no values as declared checkbox, not untyped", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Active"],
      columnValidationConditionTypes: { 0: "BOOLEAN" },
    });

    expect(valueTitles(operator, 1)).toEqual(["checkbox"]);
    expect(operator.untypedColumnsSummary()).toBeUndefined();
  });

  it("treats BOOLEAN validation on the first data row as declared checkbox", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Active"],
      topDataRow: [{ value: null, dataValidationConditionType: "BOOLEAN" }],
    });

    expect(valueTitles(operator, 1)).toEqual(["checkbox"]);
    expect(operator.untypedColumnsSummary()).toBeUndefined();
  });

  it("keeps a value source winning over BOOLEAN validation", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Description", "Kind"],
      topDataRow: ["Unit (base)", "Unit (base)"],
      columnTypes: { 1: "TEXT" },
      columnValidationValues: { 0: ["=Item[Kind]"] },
      columnValidationConditionTypes: { 0: "BOOLEAN" },
    });

    expect(operator.newColumnConfigs().item?.description?.valueName).toBe(
      "itemKind",
    );
    expect(operator.untypedColumnsSummary()).toBeUndefined();
  });

  it("ignores a non-BOOLEAN data validation condition for value-name declaration", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Status"],
      topDataRow: ["open"],
      columnValidationConditionTypes: { 0: "ONE_OF_LIST" },
      columnValidationValues: { 0: ["open", "closed"] },
    });

    expect(valueTitles(operator, 1)).toEqual(["string"]);
    expect(operator.untypedColumnsSummary()).toContain(
      "1 column(s) across 1 Table(s)",
    );
  });

  it("prefers the declared type over what the top data row samples to", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Amount"],
      topDataRow: ["not a number at all"],
      columnTypes: { 0: "CURRENCY" },
    });

    expect(valueTitles(operator, 1)).toEqual(["number"]);
    expect(operator.untypedColumnsSummary()).toBeUndefined();
  });

  it("prefers the declared type over an empty top data row", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Unit Cost", "Shipping Date"],
      columnTypes: { 0: "CURRENCY", 1: "DATE" },
    });

    expect(valueTitles(operator, 2)).toEqual(["number", "date"]);
    expect(operator.untypedColumnsSummary()).toBeUndefined();
  });

  it("keeps the ID value name for the ID column whatever type it declares", () => {
    const operator = syncColumnsUnderTest({
      headers: ["ID"],
      topDataRow: ["test:abc"],
      columnTypes: { 0: "TEXT" },
    });

    expect(valueTitles(operator, 1)).toEqual(["id"]);
    expect(operator.untypedColumnsSummary()).toBeUndefined();
  });

  it("keeps a value source winning over a declared dropdown type", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Description", "Kind"],
      topDataRow: ["Unit (base)", "Unit (base)"],
      columnTypes: { 0: "DROPDOWN", 1: "TEXT" },
      columnValidationValues: { 0: ["=Item[Kind]"] },
    });

    expect(operator.newColumnConfigs().item?.description?.valueName).toBe(
      "itemKind",
    );
    expect(operator.untypedColumnsSummary()).toBeUndefined();
  });

  it("falls back to the sample for a dropdown with no value source", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Widget Ref"],
      topDataRow: ["wdg:abc123"],
      columnTypes: { 0: "DROPDOWN" },
    });

    expect(valueTitles(operator, 1)).toEqual(["string"]);
    expect(operator.untypedColumnsSummary()).toContain(
      "1 column(s) across 1 Table(s)",
    );
  });

  it("falls back to the sample for a column with no declared type, and counts it", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Amount"],
      topDataRow: [42],
    });

    expect(valueTitles(operator, 1)).toEqual(["number"]);
    expect(operator.untypedColumnsSummary()).toContain(
      "1 column(s) across 1 Table(s)",
    );
  });

  it("treats a compatible empty first data row's number format as declared, not untyped", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Amount", "Shipping Date", "Shipping Time", "Shipped At"],
      topDataRow: [
        { value: null, numberFormatType: "CURRENCY" },
        { value: null, numberFormatType: "DATE" },
        { value: null, numberFormatType: "TIME" },
        { value: null, numberFormatType: "DATE_TIME" },
      ],
    });

    expect(valueTitles(operator, 4)).toEqual([
      "number",
      "date",
      "number",
      "number",
    ]);
    expect(operator.untypedColumnsSummary()).toBeUndefined();
  });

  it("keeps an incompatible text sample in a Currency-formatted column guessed string and untyped", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Notes"],
      topDataRow: [{ value: "a note", numberFormatType: "CURRENCY" }],
    });

    expect(valueTitles(operator, 1)).toEqual(["string"]);
    expect(operator.untypedColumnsSummary()).toContain(
      "1 column(s) across 1 Table(s)",
    );
  });

  it("treats a compatible Plain text format as declared string, not untyped", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Notes"],
      topDataRow: [{ value: "a note", numberFormatType: "TEXT" }],
    });

    expect(valueTitles(operator, 1)).toEqual(["string"]);
    expect(operator.untypedColumnsSummary()).toBeUndefined();
  });

  it("treats a compatible Scientific format as declared number, not untyped", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Amount"],
      topDataRow: [{ value: 1.2e3, numberFormatType: "SCIENTIFIC" }],
    });

    expect(valueTitles(operator, 1)).toEqual(["number"]);
    expect(operator.untypedColumnsSummary()).toBeUndefined();
  });

  it("keeps a number in a Plain text-formatted column guessed number and untyped", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Notes"],
      topDataRow: [{ value: 42, numberFormatType: "TEXT" }],
    });

    expect(valueTitles(operator, 1)).toEqual(["number"]);
    expect(operator.untypedColumnsSummary()).toContain(
      "1 column(s) across 1 Table(s)",
    );
  });

  it("keeps TRUE with only a number format guessed boolean and untyped", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Active"],
      topDataRow: [{ value: true, numberFormatType: "CURRENCY" }],
    });

    expect(valueTitles(operator, 1)).toEqual(["boolean"]);
    expect(operator.untypedColumnsSummary()).toContain(
      "1 column(s) across 1 Table(s)",
    );
  });

  it("declares from a compatible format when the type menu is DROPDOWN with no Value Config rule", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Amount"],
      topDataRow: [{ value: 42, numberFormatType: "NUMBER" }],
      columnTypes: { 0: "DROPDOWN" },
    });

    expect(valueTitles(operator, 1)).toEqual(["number"]);
    expect(operator.untypedColumnsSummary()).toBeUndefined();
  });

  it("falls back to text for an empty top cell with no number format", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Notes"],
    });

    expect(valueTitles(operator, 1)).toEqual(["string"]);
    expect(operator.untypedColumnsSummary()).toContain(
      "1 column(s) across 1 Table(s)",
    );
  });

  it("counts a formula column alongside the rest", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Balance"],
      topDataRow: [{ value: 42, isFormula: true }],
    });

    expect(operator.newColumnConfigs().item?.balance?.isFormula).toBe(true);
    expect(operator.untypedColumnsSummary()).toContain(
      "1 column(s) across 1 Table(s)",
    );
  });

  it("summarises how many columns in how many Tables are still untyped", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Amount", "Notes", "Moved In"],
      topDataRow: [42, "a note"],
      columnTypes: { 2: "DATE" },
    });

    expect(operator.untypedColumnsSummary()).toBe(
      "Succeeded, but 2 column(s) across 1 Table(s) are untyped, so their " +
        "value names were guessed. See the execution log for the list.",
    );
  });

  it("summarises nothing when every column declares its type", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Amount", "Notes"],
      columnTypes: { 0: "CURRENCY", 1: "TEXT" },
    });

    expect(operator.untypedColumnsSummary()).toBeUndefined();
  });
});

describe("ColumnConfigsOperator.recordUntypedColumns -> a sheet whose only data row is blank", () => {
  it("completes the sync and emits every column", () => {
    const operator = syncBlankSheetUnderTest({
      headers: ["Supplier Name", "Amount"],
      columnTypes: { 1: "CURRENCY" },
    });
    const emitted = operator.newColumnConfigs().item;

    expect(emitted?.supplierName?.valueName).toBe("string");
    expect(emitted?.amount?.valueName).toBe("number");
  });

  it("records a column as no formula when the blank row proves nothing", () => {
    const operator = syncBlankSheetUnderTest({ headers: ["Amount"] });

    expect(operator.newColumnConfigs().item?.amount?.isFormula).toBe(false);
  });

  it("notes the Tables whose guesses had no sample row behind them", () => {
    const operator = syncBlankSheetUnderTest({ headers: ["Notes"] });

    expect(operator.untypedColumnsSummary()).toBe(
      "Succeeded, but 1 column(s) across 1 Table(s) are untyped, so their " +
        "value names were guessed. See the execution log for the list. " +
        "On 1 of those Table(s) the top data row was blank, so the guess had " +
        'no sample behind it: "Item".',
    );
  });

  it("stays silent for a blank sheet whose columns all declare their type", () => {
    const operator = syncBlankSheetUnderTest({
      headers: ["Amount"],
      columnTypes: { 0: "CURRENCY" },
    });

    expect(operator.untypedColumnsSummary()).toBeUndefined();
  });

  it("syncs the columns of a brand-new sheet whose data row never existed", () => {
    const operator = syncBlankSheetUnderTest({
      headers: ["Supplier Name"],
      topDataRowAbsence: "rowsWithNoGridBlock",
    });

    expect(operator.newColumnConfigs().item?.supplierName).toMatchObject({
      header: "Supplier Name",
      valueName: "string",
    });
  });

  it("says nothing about a sample row that holds data", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Notes"],
      topDataRow: ["a note"],
    });

    expect(operator.untypedColumnsSummary()).not.toContain("no sample");
  });
});

describe("ColumnConfigsOperator.addMissingColumnIds", () => {
  it("adds a missing column ID, under the previous configs' prefix, only on a Table the configs grant", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: itemSheetGid,
          title: "Item",
          rows: buildGridRows({
            0: [""],
            3: ["Amount"],
            4: [42],
          }),
          table: { tableId: itemTableId, name: "Item", endRowIndex: 5 },
        },
        {
          sheetId: ungrantedGid,
          title: "Ghost",
          rows: buildGridRows({ 0: [""], 3: ["Amount"] }),
          table: { tableId: "ghost-table", name: "Ghost", endRowIndex: 5 },
        },
      ],
    });

    const operator = syncedOperator();

    expect(
      operator.ss.raw.table(itemTableId).headRow("columnId").valueOrEmpty(0),
    ).toMatch(/^c:itm:/);
    expect(operator.tableConfigOperator.tableIdsApiAccesses()).not.toContain(
      "ghost-table",
    );
  });

  it("adds a missing column ID when the columnId row has never had any grid data set", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: itemSheetGid,
          title: "Item",
          rows: buildGridRows({
            3: ["Amount"],
            4: [42],
          }),
          // Row 0 has never had a column ID written, so Google's real API
          // omits it entirely from the fetch response rather than
          // returning empty cells for it.
          rowsWithNoGridData: [0],
          table: { tableId: itemTableId, name: "Item", endRowIndex: 5 },
        },
      ],
    });

    const operator = syncedOperator();

    expect(
      operator.ss.raw.table(itemTableId).headRow("columnId").valueOrEmpty(0),
    ).not.toBe("");
  });
});
