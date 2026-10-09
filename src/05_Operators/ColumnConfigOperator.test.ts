import type { StrictOmit } from "@byronbroughten/utils/obj";
import { beforeEach, describe, expect, it } from "vitest";

import { installedConfigs } from "../01_SpreadsheetSchema/configReaders/configRegister";
import { configSheetFloorSeed } from "../01_SpreadsheetSchema/configReaders/configSheetFloorSeed";
import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  type FakeCell,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { tableIdOnTab } from "../testSupport/fakeTableConfigSheet";
import { ColumnConfigOperator } from "./ColumnConfigOperator";

const { columnConfigs } = installedConfigs();

// Real committed columnId strings, so fixtures stay honest to what the
// production code actually resolves column names through.
const tc = columnConfigs.tableConfig;
const cc = columnConfigs.columnConfig;
const tableConfigGid = 210603630;
const columnConfigGid = 2034522667;
const widgetGid = 999001;
const newSheetGid = 999002;
const unresolvableGid = 424242;

const testSheetGid = getTableTraitByName("item", "sheetGid");
const tableConfigTableId = tableIdOnTab(tableConfigGid);
const columnConfigTableId = tableIdOnTab(columnConfigGid);
const widgetTableId = tableIdOnTab(widgetGid);
const newSheetTableId = tableIdOnTab(newSheetGid);
const unresolvableTableId = tableIdOnTab(unresolvableGid);
const testTableId = tableIdOnTab(testSheetGid);

// The floor Tables carry their generated names, which give their Table keys.
const tableConfigTable = {
  tableId: tableConfigTableId,
  name: configSheetFloorSeed.tableConfig.liveTableName,
};
const columnConfigTable = {
  tableId: columnConfigTableId,
  name: configSheetFloorSeed.columnConfig.liveTableName,
};

const tableConfigColumnIdRow = [
  tc.tableId.columnId,
  tc.tableName.columnId,
  tc.sheetTitle.columnId,
  tc.letApiAccess.columnId,
];
const tableConfigHeaderRow = [
  tc.tableId.header,
  tc.tableName.header,
  tc.sheetTitle.header,
  tc.letApiAccess.header,
];
const tableConfigColumnTypes = {
  0: "TEXT",
  1: "TEXT",
  2: "TEXT",
  3: "BOOLEAN",
} as const;

const columnConfigColumnIdRow = [
  cc.tableId.columnId,
  cc.columnId.columnId,
  cc.tableName.columnId,
  cc.header.columnId,
  cc.emptyValueAllowed.columnId,
];
const columnConfigHeaderRow = [
  cc.tableId.header,
  cc.columnId.header,
  cc.tableName.header,
  cc.header.header,
  cc.emptyValueAllowed.header,
];
const columnConfigColumnTypes = {
  0: "TEXT",
  1: "TEXT",
  2: "TEXT",
  3: "TEXT",
  4: "BOOLEAN",
} as const;

const freshlyAppendedRowMissingHeaderAndValueName = [
  widgetTableId,
  "c:wdg:ddd",
  "Widget",
  "",
];
const rowReferencingUnresolvableTable = [
  unresolvableTableId,
  "c:???:eee",
  "",
  "Orphan Field",
];

beforeEach(() => {
  stubLogger();
});

// Syncs Table Config (so tableId -> Table key resolves for Widget/Brand
// New Sheet, via auto-appended rows) and fetches whatever Column Config
// rows the caller seeded — without running the full append/prune column-ID
// lifecycle, keeping these tests focused on toFileSource's own read/skip/
// throw logic rather than re-testing the pre-existing lifecycle.
function initSyncedColumnConfigOperator(): ColumnConfigOperator {
  const columnConfigOperator = ColumnConfigOperator.init();
  const tableConfigOperator = columnConfigOperator.tableConfigOperator;
  columnConfigOperator.ss.fetchAllSheetProperties();
  tableConfigOperator.table.prepFetchColumnsFull("letApiAccess");
  tableConfigOperator.prepFetchForSync();
  columnConfigOperator.table.prepFetchColumnsFull(
    "tableId",
    "columnId",
    "header",
    "emptyValueAllowed",
  );
  columnConfigOperator.ss.fetchAllPrepped({ skipFetchingProperties: true });
  tableConfigOperator.syncToSpreadsheet();
  columnConfigOperator.fetchAfterTableConfigSynced();
  return columnConfigOperator;
}

function stubGroupedColumnConfigSheets(): void {
  stubSheetsService({
    sheets: [
      {
        sheetId: tableConfigGid,
        title: "Table Config",
        rows: buildGridRows({
          0: tableConfigColumnIdRow,
          3: tableConfigHeaderRow,
          4: [widgetTableId, "", "Widget", true],
          5: [newSheetTableId, "", "Brand New Sheet", true],
        }),
        table: { ...tableConfigTable, endRowIndex: 6 },
      },
      {
        sheetId: columnConfigGid,
        title: "Column Config",
        rows: buildGridRows({
          0: columnConfigColumnIdRow,
          3: columnConfigHeaderRow,
          4: [widgetTableId, "c:wdg:aaa", "Widget", "Unit Price"],
          5: [widgetTableId, "c:wdg:bbb", "Widget", "Notes"],
          6: [newSheetTableId, "c:999002:ccc", "Brand New Sheet", "Some Field"],
        }),
        table: { ...columnConfigTable, endRowIndex: 7 },
      },
      {
        sheetId: widgetGid,
        title: "Widget",
        rows: buildGridRows({
          0: ["c:wdg:aaa", "c:wdg:bbb"],
          3: ["Unit Price", "Notes"],
          4: [42, "a note"],
        }),
        table: { name: "Widget", endRowIndex: 5 },
      },
      {
        sheetId: newSheetGid,
        title: "Brand New Sheet",
        rows: buildGridRows({
          0: ["c:999002:ccc"],
          3: ["Some Field"],
          4: ["x"],
        }),
        table: { name: "Brand New Sheet", endRowIndex: 5 },
      },
    ],
  });
}

// Mirrors ConfigCoordinator.syncAndFlushConfigSheets's own sequence (see
// docs/generated-data.md on why Table Config and Column Config sync together),
// stopping short of the final batchUpdateGSheets flush these tests don't
// need.
function syncColumnConfigOperator(operator: ColumnConfigOperator): void {
  operator.ss.fetchAllSheetProperties();
  const tableConfigOperator = operator.tableConfigOperator;
  tableConfigOperator.prepFetchForSync();
  operator.prepFetchWithTableConfig();
  operator.ss.fetchAllPrepped({ skipFetchingProperties: true });
  tableConfigOperator.syncToSpreadsheet();
  operator.fetchAfterTableConfigSynced();
  operator.syncToSpreadsheet();
}

describe("ColumnConfigOperator.newColumnConfigs / toFileSource", () => {
  it("groups columns by Table key", () => {
    stubGroupedColumnConfigSheets();

    const entries = initSyncedColumnConfigOperator().newColumnConfigs();

    expect(entries.widget).toEqual({
      unitPrice: {
        columnId: "c:wdg:aaa",
        valueName: "number",
        header: "Unit Price",
        isFormula: false,
        emptyValueAllowed: false,
        customDefaultValue: null,
      },
      notes: {
        columnId: "c:wdg:bbb",
        valueName: "string",
        header: "Notes",
        isFormula: false,
        emptyValueAllowed: false,
        customDefaultValue: null,
      },
    });
    expect(entries.brandNewSheet).toEqual({
      someField: {
        columnId: "c:999002:ccc",
        valueName: "string",
        header: "Some Field",
        isFormula: false,
        emptyValueAllowed: false,
        customDefaultValue: null,
      },
    });
  });

  it("emits one labeled column-config record per line", () => {
    stubGroupedColumnConfigSheets();

    const sourceLines = initSyncedColumnConfigOperator()
      .toFileSource("../makeConfigs")
      .split("\n");

    expect(sourceLines).toContain(
      '    "unitPrice": { "columnId": "c:wdg:aaa", "header": "Unit Price", "valueName": "number", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },',
    );
    expect(
      sourceLines.find((line) => line.includes('"widget":')),
    ).not.toContain('"columnId"');
    sourceLines
      .filter((line) => line.includes('"columnId"'))
      .forEach((line) => {
        expect(line).toContain('"header"');
        expect(line).toContain('"valueName"');
        expect(line).toContain('"isFormula"');
        expect(line).toContain('"emptyValueAllowed"');
        expect(line).toContain('"customDefaultValue"');
      });
  });

  it("emits the empty-value-allowed trait each column's own box declares", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({
            0: tableConfigColumnIdRow,
            3: tableConfigHeaderRow,
            4: [widgetTableId, "", "Widget", true],
          }),
          table: { ...tableConfigTable, endRowIndex: 5 },
        },
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
            4: [widgetTableId, "c:wdg:aaa", "Widget", "Unit Price", true],
            5: [widgetTableId, "c:wdg:bbb", "Widget", "Notes", false],
          }),
          table: { ...columnConfigTable, endRowIndex: 6 },
        },
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({
            0: ["c:wdg:aaa", "c:wdg:bbb"],
            3: ["Unit Price", "Notes"],
            4: [42, "a note"],
          }),
          table: { name: "Widget", endRowIndex: 5 },
        },
      ],
    });

    const entries = initSyncedColumnConfigOperator().newColumnConfigs();

    expect(entries.widget?.unitPrice?.emptyValueAllowed).toBe(true);
    expect(entries.widget?.notes?.emptyValueAllowed).toBe(false);
  });

  it("throws when a row is missing its header or value name", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({ 0: tableConfigColumnIdRow }),
          table: { ...tableConfigTable, endRowIndex: 5 },
        },
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
            4: freshlyAppendedRowMissingHeaderAndValueName,
          }),
          table: { ...columnConfigTable, endRowIndex: 5 },
        },
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({ 3: [] }),
        },
      ],
    });

    expect(() => initSyncedColumnConfigOperator().newColumnConfigs()).toThrow(
      /is empty/,
    );
  });

  it("throws when a row references a Table ID with no ticked Table Config row", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({ 0: tableConfigColumnIdRow }),
          table: { ...tableConfigTable, endRowIndex: 5 },
        },
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
            4: rowReferencingUnresolvableTable,
          }),
          table: { ...columnConfigTable, endRowIndex: 5 },
        },
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({ 3: [] }),
        },
      ],
    });

    expect(() => initSyncedColumnConfigOperator().newColumnConfigs()).toThrow(
      /has no ticked row in Table Config/,
    );
  });

  it("throws when two headers in the same Table camelCase to the same column name", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({
            0: tableConfigColumnIdRow,
            3: tableConfigHeaderRow,
            4: [widgetTableId, "", "Widget", true],
          }),
          table: { ...tableConfigTable, endRowIndex: 5 },
        },
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
            4: [widgetTableId, "c:wdg:aaa", "Widget", "Unit Price"],
            5: [widgetTableId, "c:wdg:bbb", "Widget", "Unit  Price"],
          }),
          table: { ...columnConfigTable, endRowIndex: 6 },
        },
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({
            0: ["c:wdg:aaa", "c:wdg:bbb"],
            3: ["Unit Price", "Unit  Price"],
            4: [1, 2],
          }),
          table: { name: "Widget", endRowIndex: 5 },
        },
      ],
    });

    expect(() => initSyncedColumnConfigOperator().newColumnConfigs()).toThrow(
      /duplicate column name "unitPrice"/,
    );
  });
});

describe("ColumnConfigOperator.syncToSpreadsheet -> _updateProgrammaticValues", () => {
  const testTableConfigRowWithApiAccess = [testTableId, "", "Item", true];

  function seedTableConfigFixture() {
    return {
      sheetId: tableConfigGid,
      title: "Table Config",
      rows: buildGridRows({
        0: tableConfigColumnIdRow,
        3: tableConfigHeaderRow,
        4: testTableConfigRowWithApiAccess,
      }),
      table: {
        ...tableConfigTable,
        endRowIndex: 5,
        columnTypes: tableConfigColumnTypes,
      },
    };
  }

  it("corrects Table name and header, keying emitted columns by Table key", () => {
    stubSheetsService({
      sheets: [
        seedTableConfigFixture(),
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
            4: [
              testTableId,
              "c:itm:corr01",
              "Stale Name",
              "Stale Header",
              true,
            ],
            5: [testTableId, "c:itm:corr02", "Item", "ID"],
          }),
          table: { ...columnConfigTable, endRowIndex: 6 },
        },
        {
          sheetId: testSheetGid,
          title: "Item tab",
          rows: buildGridRows({
            0: ["c:itm:corr01", "c:itm:corr02"],
            3: ["Amount", "ID"],
            4: [42, "xyz"],
          }),
          table: { name: "Item", endRowIndex: 5 },
        },
      ],
    });

    const operator = ColumnConfigOperator.init();
    syncColumnConfigOperator(operator);
    const identity = operator.table.columns("tableName", "header");
    const emitted = operator.newColumnConfigs().item;

    expect(identity.tableName.value(0)).toBe("Item");
    expect(identity.header.value(0)).toBe("Amount");
    expect(emitted?.amount).toMatchObject({
      valueName: "number",
      isFormula: false,
      emptyValueAllowed: true,
    });
    expect(operator.table.column("emptyValueAllowed").value(0)).toBe(true);

    expect(identity.tableName.value(1)).toBe("Item");
    expect(identity.header.value(1)).toBe("ID");
    expect(emitted?.id).toMatchObject({
      valueName: "id",
      isFormula: false,
    });
  });

  it("fills in a row whose identity cells have never been filled in", () => {
    stubSheetsService({
      sheets: [
        seedTableConfigFixture(),
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
            4: [testTableId, "c:itm:corr05", null, null],
          }),
          table: { ...columnConfigTable, endRowIndex: 5 },
        },
        {
          sheetId: testSheetGid,
          title: "Item",
          rows: buildGridRows({
            0: ["c:itm:corr05"],
            3: ["Amount"],
            4: [42],
          }),
          table: { name: "Item", endRowIndex: 5 },
        },
      ],
    });

    const operator = ColumnConfigOperator.init();
    syncColumnConfigOperator(operator);
    const identity = operator.table.columns("tableName", "header");

    expect(identity.tableName.value(0)).toBe("Item");
    expect(identity.header.value(0)).toBe("Amount");
    expect(operator.newColumnConfigs().item?.amount).toMatchObject({
      valueName: "number",
      isFormula: false,
    });
  });

  it("writes a floor column's ticked Empty value allowed back to FALSE and leaves a business column's tick", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({
            0: tableConfigColumnIdRow,
            3: tableConfigHeaderRow,
            4: [testTableId, "", "Item", true],
            5: [tableConfigTableId, "", "Table Config", true],
          }),
          table: {
            ...tableConfigTable,
            endRowIndex: 6,
            endColumnIndex: 4,
            columnTypes: tableConfigColumnTypes,
          },
        },
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
            4: [testTableId, "c:itm:corr01", "Item", "Amount", true],
            5: [
              tableConfigTableId,
              tc.tableId.columnId,
              "Table Config",
              tc.tableId.header,
              true,
            ],
          }),
          table: {
            ...columnConfigTable,
            endRowIndex: 6,
            columnTypes: columnConfigColumnTypes,
          },
        },
        {
          sheetId: testSheetGid,
          title: "Item",
          rows: buildGridRows({
            0: ["c:itm:corr01"],
            3: ["Amount"],
            4: [42],
          }),
          table: { name: "Item", endRowIndex: 5 },
        },
      ],
    });

    const operator = ColumnConfigOperator.init();
    syncColumnConfigOperator(operator);
    const col = operator.table.columns("columnId", "emptyValueAllowed");
    const businessRow = operator.table.workingRowIndexesWithData.find(
      (rowIndex) => col.columnId.value(rowIndex) === "c:itm:corr01",
    );
    const floorRow = operator.table.workingRowIndexesWithData.find(
      (rowIndex) => col.columnId.value(rowIndex) === tc.tableId.columnId,
    );

    expect(businessRow).toBeDefined();
    expect(floorRow).toBeDefined();
    if (businessRow === undefined || floorRow === undefined) {
      throw new Error("Expected business and floor Column Config rows.");
    }
    expect(col.emptyValueAllowed.value(businessRow)).toBe(true);
    expect(col.emptyValueAllowed.value(floorRow)).toBe(false);
    expect(operator.declaredCellReport()).toContain(
      `Column Config · Empty value allowed · tableConfig · ${tc.tableId.header} → FALSE`,
    );
  });

  it("detects a named valueConfig from the column's live data-validation formula", () => {
    stubSheetsService({
      sheets: [
        seedTableConfigFixture(),
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
            4: [testTableId, "c:itm:corr03", "Item", "Description"],
          }),
          table: { ...columnConfigTable, endRowIndex: 5 },
        },
        {
          sheetId: testSheetGid,
          title: "Item",
          rows: buildGridRows({
            0: ["c:itm:corr03"],
            3: ["Description"],
            4: ["Unit (base)"],
          }),
          table: {
            name: "Item",
            endRowIndex: 5,
            columnValidationValues: {
              0: ["=valueConfig[Transaction Description]"],
            },
          },
        },
      ],
    });

    const operator = ColumnConfigOperator.init();
    syncColumnConfigOperator(operator);
    const identity = operator.table.columns("tableName", "header");

    expect(valueTitles(operator, 1)).toEqual(["Transaction Description"]);
    expect(operator.newColumnConfigs().item?.description?.valueName).toBe(
      "transactionDescription",
    );
    expect(identity.tableName.value(0)).toBe("Item");
    expect(identity.header.value(0)).toBe("Description");
  });

  it("detects a live formula and a date-formatted number", () => {
    stubSheetsService({
      sheets: [
        seedTableConfigFixture(),
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
            4: [testTableId, "c:itm:corr04", "Item", "Due-by Date"],
          }),
          table: { ...columnConfigTable, endRowIndex: 5 },
        },
        {
          sheetId: testSheetGid,
          title: "Item",
          rows: buildGridRows({
            0: ["c:itm:corr04"],
            3: ["Due-by Date"],
            4: [{ value: 45000, isFormula: true, numberFormatType: "DATE" }],
          }),
          table: { name: "Item", endRowIndex: 5 },
        },
      ],
    });

    const operator = ColumnConfigOperator.init();
    syncColumnConfigOperator(operator);

    expect(operator.newColumnConfigs().item?.dueByDate).toMatchObject({
      valueName: "date",
      isFormula: true,
    });
  });
});

describe("ColumnConfigOperator.syncToSpreadsheet -> _appendColumnRows", () => {
  const propertyColumnConfigRow = [
    widgetTableId,
    "c:wdg:aaa",
    "Widget",
    "Unit Price",
  ];
  const newSheetColumnConfigRow = [
    newSheetTableId,
    "c:wdg:aaa",
    "Brand New Sheet",
    "Unit Price",
  ];

  function seedDuplicatedColumnIdOnTwoSheets(columnConfigRows: unknown[][]) {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({
            0: tableConfigColumnIdRow,
            3: tableConfigHeaderRow,
            4: [widgetTableId, "", "Widget", true],
            5: [newSheetTableId, "", "Brand New Sheet", true],
          }),
          table: {
            ...tableConfigTable,
            endRowIndex: 6,
            columnTypes: tableConfigColumnTypes,
          },
        },
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
            ...Object.fromEntries(
              columnConfigRows.map((row, index) => [4 + index, row]),
            ),
          }),
          table: {
            ...columnConfigTable,
            endRowIndex: 4 + columnConfigRows.length,
            columnTypes: columnConfigColumnTypes,
          },
        },
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({
            0: ["c:wdg:aaa"],
            3: ["Unit Price"],
            4: [42],
          }),
          table: { name: "Widget", endRowIndex: 5 },
        },
        {
          sheetId: newSheetGid,
          title: "Brand New Sheet",
          rows: buildGridRows({
            0: ["c:wdg:aaa"],
            3: ["Unit Price"],
            4: [42],
          }),
          table: { name: "Brand New Sheet", endRowIndex: 5 },
        },
      ],
    });
  }

  function identitiesOnTheTwoTables(operator: ColumnConfigOperator): string[] {
    const col = operator.table.columns("tableId", "columnId");
    return operator.table.workingRowIndexesWithData
      .map((rowIndex) => [
        col.tableId.value(rowIndex),
        col.columnId.value(rowIndex),
      ])
      .filter(
        ([tableId]) => tableId === widgetTableId || tableId === newSheetTableId,
      )
      .map(([tableId, columnId]) => `${tableId}:${columnId}`);
  }

  it("appends a row for a column whose ID already has a row under another Table", () => {
    seedDuplicatedColumnIdOnTwoSheets([propertyColumnConfigRow]);

    const operator = ColumnConfigOperator.init();
    syncColumnConfigOperator(operator);

    expect(identitiesOnTheTwoTables(operator)).toEqual([
      `${widgetTableId}:c:wdg:aaa`,
      `${newSheetTableId}:c:wdg:aaa`,
    ]);
  });

  it("appends no row for a column whose identity already has one", () => {
    seedDuplicatedColumnIdOnTwoSheets([
      propertyColumnConfigRow,
      newSheetColumnConfigRow,
    ]);

    const operator = ColumnConfigOperator.init();
    syncColumnConfigOperator(operator);

    expect(identitiesOnTheTwoTables(operator)).toEqual([
      `${widgetTableId}:c:wdg:aaa`,
      `${newSheetTableId}:c:wdg:aaa`,
    ]);
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
}: ColumnsUnderTest): ColumnConfigOperator {
  const columnIds = headers.map((_, index) => `c:itm:col${index}`);
  const columnConfigRows: Record<number, FakeCell[]> = {
    0: columnConfigColumnIdRow,
    3: columnConfigHeaderRow,
  };
  columnIds.forEach((columnId, index) => {
    columnConfigRows[4 + index] = [testTableId, columnId, "Item", ""];
  });
  stubSheetsService({
    sheets: [
      {
        sheetId: tableConfigGid,
        title: "Table Config",
        rows: buildGridRows({
          0: tableConfigColumnIdRow,
          3: tableConfigHeaderRow,
          4: [testTableId, "", "Item", true],
        }),
        table: {
          ...tableConfigTable,
          endRowIndex: 5,
          columnTypes: tableConfigColumnTypes,
        },
      },
      {
        sheetId: columnConfigGid,
        title: "Column Config",
        rows: buildGridRows(columnConfigRows),
        table: {
          ...columnConfigTable,
          endRowIndex: Math.max(5, 4 + columnIds.length),
          columnTypes: columnConfigColumnTypes,
        },
      },
      {
        sheetId: testSheetGid,
        title: "Item",
        rows: buildGridRows({ 0: columnIds, 3: headers, 4: topDataRow }),
        ...(topDataRowAbsence ? { [topDataRowAbsence]: [4] } : {}),
        table: {
          name: "Item",
          endRowIndex: 5,
          columnTypes,
          columnValidationValues,
          columnValidationConditionTypes,
        },
      },
    ],
  });
  const operator = ColumnConfigOperator.init();
  syncColumnConfigOperator(operator);
  return operator;
}

function valueTitles(operator: ColumnConfigOperator, count: number) {
  const col = operator.table.columns("tableId", "columnId");
  const titles = operator.table.workingRowIndexesWithData.flatMap(
    (rowIndex) => {
      if (col.tableId.valueOrEmpty(rowIndex) !== testTableId) return [];
      return [
        operator.ss.raw
          .table(testTableId)
          .profile.columnById(col.columnId.value(rowIndex))
          .valueTitle(),
      ];
    },
  );
  expect(titles).toHaveLength(count);
  return titles;
}

// The reported bug's shape: a sheet emptied by a run that consumed its input.
function syncBlankSheetUnderTest(
  props: StrictOmit<ColumnsUnderTest, "topDataRow">,
): ColumnConfigOperator {
  return syncColumnsUnderTest({
    topDataRowAbsence: "rowsWithNoGridData",
    ...props,
  });
}

describe("ColumnConfigOperator.syncToSpreadsheet -> declared column types", () => {
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

  it("keeps a Value Config rule winning over BOOLEAN validation", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Description"],
      topDataRow: ["Unit (base)"],
      columnValidationValues: {
        0: ["=valueConfig[Transaction Description]"],
      },
      columnValidationConditionTypes: { 0: "BOOLEAN" },
    });

    expect(valueTitles(operator, 1)).toEqual(["Transaction Description"]);
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

  it("keeps a Value Config validation rule winning over a declared dropdown type", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Description"],
      topDataRow: ["Unit (base)"],
      columnTypes: { 0: "DROPDOWN" },
      columnValidationValues: {
        0: ["=valueConfig[Transaction Description]"],
      },
    });

    expect(valueTitles(operator, 1)).toEqual(["Transaction Description"]);
    expect(operator.untypedColumnsSummary()).toBeUndefined();
  });

  it("falls back to the sample for a dropdown with no Value Config rule", () => {
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

describe("ColumnConfigOperator.syncToSpreadsheet -> a sheet whose only data row is blank", () => {
  it("completes the sync and fills in every identity cell", () => {
    const operator = syncBlankSheetUnderTest({
      headers: ["Supplier Name", "Amount"],
      columnTypes: { 1: "CURRENCY" },
    });
    const col = operator.table.columns("tableName", "header");
    const emitted = operator.newColumnConfigs().item;

    expect(col.tableName.value(0)).toBe("Item");
    expect(col.header.value(0)).toBe("Supplier Name");
    expect(emitted?.supplierName?.valueName).toBe("string");
    expect(col.header.value(1)).toBe("Amount");
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
    const col = operator.table.columns("header");

    expect(col.header.value(0)).toBe("Supplier Name");
    expect(operator.newColumnConfigs().item?.supplierName?.valueName).toBe(
      "string",
    );
  });

  it("says nothing about a sample row that holds data", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Notes"],
      topDataRow: ["a note"],
    });

    expect(operator.untypedColumnsSummary()).not.toContain("no sample");
  });
});

describe("ColumnConfigOperator.syncToSpreadsheet -> _addMissingColumnIds", () => {
  it("adds a missing column ID only for the letApiAccess=true sheet, skipping the letApiAccess=false one without fetching it", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({
            0: tableConfigColumnIdRow,
            3: tableConfigHeaderRow,
            // Api-access sheet: gets a missing column ID filled in.
            4: [testTableId, "", "Item", true],

            5: [unresolvableTableId, "", "Ghost", false],
          }),
          table: { ...tableConfigTable, endRowIndex: 6 },
        },
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
          }),
          table: { ...columnConfigTable, endRowIndex: 5 },
        },
        {
          sheetId: testSheetGid,
          title: "Item",
          rows: buildGridRows({
            0: [""],
            3: ["Amount"],
            4: [42],
          }),
          table: { name: "Item", endRowIndex: 5 },
        },
      ],
    });

    const operator = ColumnConfigOperator.init();

    expect(() => syncColumnConfigOperator(operator)).not.toThrow();

    const colIdRow = operator.ss.raw.table(testTableId).headRow("columnId");
    expect(colIdRow.valueOrEmpty(0)).not.toBe("");
  });

  it("adds a missing column ID when the columnId row has never had any grid data set", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({
            0: tableConfigColumnIdRow,
            3: tableConfigHeaderRow,
            4: [testTableId, "", "Item", true],
          }),
          table: { ...tableConfigTable, endRowIndex: 5 },
        },
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
          }),
          table: { ...columnConfigTable, endRowIndex: 5 },
        },
        {
          sheetId: testSheetGid,
          title: "Item",
          rows: buildGridRows({
            3: ["Amount"],
            4: [42],
          }),
          // Row 0 has never had a column ID written, so Google's real API
          // omits it entirely from the fetch response rather than
          // returning empty cells for it.
          rowsWithNoGridData: [0],
          table: { name: "Item", endRowIndex: 5 },
        },
      ],
    });

    const operator = ColumnConfigOperator.init();

    expect(() => syncColumnConfigOperator(operator)).not.toThrow();

    const colIdRow = operator.ss.raw.table(testTableId).headRow("columnId");
    expect(colIdRow.valueOrEmpty(0)).not.toBe("");
  });
});

describe("ColumnConfigOperator.syncToSpreadsheet -> _pruneColumnRows", () => {
  function seedColumnConfigDescribingItselfBelowABlankTopDataRow() {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({
            0: tableConfigColumnIdRow,
            3: tableConfigHeaderRow,
            4: [columnConfigTableId, "", "Column Config", true],
          }),
          table: {
            ...tableConfigTable,
            endRowIndex: 5,
            endColumnIndex: 4,
            columnTypes: tableConfigColumnTypes,
          },
        },
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
            4: [],
            5: [
              columnConfigTableId,
              cc.tableId.columnId,
              "Stale Name",
              "Stale Header",
            ],
          }),
          table: {
            ...columnConfigTable,
            endRowIndex: 6,
            columnTypes: columnConfigColumnTypes,
          },
        },
      ],
    });
  }

  // Column Config's own Let api access starts off; the correction pass ticks
  // it, so its floor columns stay and a stale business row is pruned.
  function seedColumnConfigWithAStaleBusinessRow() {
    stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({
            0: tableConfigColumnIdRow,
            3: tableConfigHeaderRow,
            4: [columnConfigTableId, "", "Column Config", false],
          }),
          table: {
            ...tableConfigTable,
            endRowIndex: 5,
            endColumnIndex: 4,
            columnTypes: tableConfigColumnTypes,
          },
        },
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
            4: [unresolvableTableId, "c:???:eee", "Ghost", "Orphan Field"],
            5: [columnConfigTableId, cc.tableId.columnId, "columnConfig"],
          }),
          table: {
            ...columnConfigTable,
            endRowIndex: 6,
            columnTypes: columnConfigColumnTypes,
          },
        },
      ],
    });
  }

  it("prunes a stale business row and keeps floor-column rows", () => {
    seedColumnConfigWithAStaleBusinessRow();

    const operator = ColumnConfigOperator.init();

    expect(() => syncColumnConfigOperator(operator)).not.toThrow();
    expect(operator.table.column("columnId").hasValue("c:???:eee")).toBe(false);
    expect(operator.table.column("tableId").hasValue(columnConfigTableId)).toBe(
      true,
    );
    expect(operator.newColumnConfigs().columnConfig).toBeDefined();
  });

  it("still resolves programmatic values for a sheet whose own top data row it pruned", () => {
    seedColumnConfigDescribingItselfBelowABlankTopDataRow();

    const operator = ColumnConfigOperator.init();
    syncColumnConfigOperator(operator);
    const col = operator.table.columns("tableName", "header");
    const emitted = operator.newColumnConfigs().columnConfig;

    expect(operator.table.workingRowIndexes).not.toContain(0);
    expect(col.tableName.value(1)).toBe("columnConfig");
    expect(col.header.value(1)).toBe("Table ID");
    expect(emitted?.tableId).toMatchObject({
      valueName: "string",
      isFormula: false,
    });
  });
});
