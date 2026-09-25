import { beforeEach, describe, expect, it } from "vitest";

import { installedConfigs } from "../01_SpreadsheetSchema/configRegister";
import { getSheetTraitByName } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  type FakeCell,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import type { StrictOmit } from "../utils/Obj";
import { ColumnConfigOperator } from "./ColumnConfigOperator";

const { columnConfigs } = installedConfigs();

// Real committed columnId strings, so fixtures stay honest to what the
// production code actually resolves column names through.
const sc = columnConfigs.sheetConfig;
const cc = columnConfigs.columnConfig;
const sheetConfigGid = 210603630;
const columnConfigGid = 2034522667;
const widgetGid = 999001;
const newSheetGid = 999002;
const unresolvableGid = 424242;

const testSheetGid = getSheetTraitByName("item", "sheetGid");

const sheetConfigColumnIdRow = [
  sc.sheetGid.columnId,
  sc.sheetTitle.columnId,
  sc.letApiAccess.columnId,
];
const sheetConfigHeaderRow = [
  sc.sheetGid.header,
  sc.sheetTitle.header,
  sc.letApiAccess.header,
];
const sheetConfigColumnTypes = {
  0: "DOUBLE",
  1: "TEXT",
  2: "BOOLEAN",
} as const;
const columnConfigColumnIdRow = [
  cc.sheetGid.columnId,
  cc.columnId.columnId,
  cc.sheetTitle.columnId,
  cc.header.columnId,
  cc.emptyValueAllowed.columnId,
];
const columnConfigHeaderRow = [
  cc.sheetGid.header,
  cc.columnId.header,
  cc.sheetTitle.header,
  cc.header.header,
  cc.emptyValueAllowed.header,
];
const columnConfigColumnTypes = {
  0: "DOUBLE",
  1: "TEXT",
  2: "TEXT",
  3: "TEXT",
  4: "BOOLEAN",
} as const;

const freshlyAppendedRowMissingHeaderAndValueName = [
  widgetGid,
  "c:wdg:ddd",
  "Widget",
  "",
];
const rowReferencingUnresolvableSheet = [
  unresolvableGid,
  "c:???:eee",
  "",
  "Orphan Field",
];

beforeEach(() => {
  stubLogger();
});

// Syncs Sheet Config (so sheetGid -> sheetName resolves for Widget/Brand
// New Sheet, via auto-appended rows) and fetches whatever Column Config
// rows the caller seeded — without running the full append/prune column-ID
// lifecycle, keeping these tests focused on toFileSource's own read/skip/
// throw logic rather than re-testing the pre-existing lifecycle.
function initSyncedColumnConfigOperator(): ColumnConfigOperator {
  const columnConfigOperator = ColumnConfigOperator.init();
  const sheetConfigOperator = columnConfigOperator.sheetConfigOperator;
  sheetConfigOperator.sheet.prepFetchColumnsFull("letApiAccess");
  sheetConfigOperator.prepFetchForSync();
  columnConfigOperator.sheet.prepFetchColumnsFull(
    "sheetGid",
    "columnId",
    "header",
    "emptyValueAllowed",
  );
  columnConfigOperator.ss.fetchAllPrepped();
  sheetConfigOperator.syncToSpreadsheet();
  columnConfigOperator.fetchAfterSheetConfigSynced();
  return columnConfigOperator;
}

function stubGroupedColumnConfigSheets(): void {
  stubSheetsService({
    sheets: [
      {
        sheetId: sheetConfigGid,
        title: "Sheet Config",
        rows: buildGridRows({
          0: sheetConfigColumnIdRow,
          3: sheetConfigHeaderRow,
          4: [widgetGid, "Widget", true, ""],
          5: [newSheetGid, "Brand New Sheet", true, ""],
        }),
        table: { endRowIndex: 6 },
      },
      {
        sheetId: columnConfigGid,
        title: "Column Config",
        rows: buildGridRows({
          0: columnConfigColumnIdRow,
          3: columnConfigHeaderRow,
          4: [widgetGid, "c:wdg:aaa", "Widget", "Unit Price"],
          5: [widgetGid, "c:wdg:bbb", "Widget", "Notes"],
          6: [newSheetGid, "c:999002:ccc", "Brand New Sheet", "Some Field"],
        }),
        table: { endRowIndex: 7 },
      },
      {
        sheetId: widgetGid,
        title: "Widget",
        rows: buildGridRows({
          0: ["c:wdg:aaa", "c:wdg:bbb"],
          3: ["Unit Price", "Notes"],
          4: [42, "a note"],
        }),
        table: { endRowIndex: 5 },
      },
      {
        sheetId: newSheetGid,
        title: "Brand New Sheet",
        rows: buildGridRows({
          0: ["c:999002:ccc"],
          3: ["Some Field"],
          4: ["x"],
        }),
        table: { endRowIndex: 5 },
      },
    ],
  });
}

// Mirrors ConfigCoordinator.syncAndFlushConfigSheets's own sequence (see
// docs/generated-data.md on why Sheet Config and Column Config sync together),
// stopping short of the final batchUpdateGSheets flush these tests don't
// need.
function syncColumnConfigOperator(operator: ColumnConfigOperator): void {
  operator.ss.fetchAllSheetProperties();
  const sheetConfigOperator = operator.sheetConfigOperator;
  sheetConfigOperator.prepFetchForSync();
  operator.prepFetchWithSheetConfig();
  operator.ss.fetchAllPrepped({ skipFetchingProperties: true });
  sheetConfigOperator.syncToSpreadsheet();
  operator.fetchAfterSheetConfigSynced();
  operator.syncToSpreadsheet();
}

describe("ColumnConfigOperator.newColumnConfigs / toFileSource", () => {
  it("groups columns by resolved sheet name", () => {
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
          sheetId: sheetConfigGid,
          title: "Sheet Config",
          rows: buildGridRows({
            0: sheetConfigColumnIdRow,
            3: sheetConfigHeaderRow,
            4: [widgetGid, "Widget", true, ""],
          }),
          table: { endRowIndex: 5 },
        },
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
            4: [widgetGid, "c:wdg:aaa", "Widget", "Unit Price", true],
            5: [widgetGid, "c:wdg:bbb", "Widget", "Notes", false],
          }),
          table: { endRowIndex: 6 },
        },
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({
            0: ["c:wdg:aaa", "c:wdg:bbb"],
            3: ["Unit Price", "Notes"],
            4: [42, "a note"],
          }),
          table: { endRowIndex: 5 },
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
          sheetId: sheetConfigGid,
          title: "Sheet Config",
          rows: buildGridRows({ 0: sheetConfigColumnIdRow }),
          table: { endRowIndex: 5 },
        },
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
            4: freshlyAppendedRowMissingHeaderAndValueName,
          }),
          table: { endRowIndex: 5 },
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

  it("throws when a row references a sheetGid unresolvable in Sheet Config", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: sheetConfigGid,
          title: "Sheet Config",
          rows: buildGridRows({ 0: sheetConfigColumnIdRow }),
          table: { endRowIndex: 5 },
        },
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
            4: rowReferencingUnresolvableSheet,
          }),
          table: { endRowIndex: 5 },
        },
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({ 3: [] }),
        },
      ],
    });

    expect(() => initSyncedColumnConfigOperator().newColumnConfigs()).toThrow(
      /no corresponding sheet name in Sheet Config/,
    );
  });

  it("throws when two headers on the same sheet camelCase to the same column name", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: sheetConfigGid,
          title: "Sheet Config",
          rows: buildGridRows({
            0: sheetConfigColumnIdRow,
            3: sheetConfigHeaderRow,
            4: [widgetGid, "Widget", true, ""],
          }),
          table: { endRowIndex: 5 },
        },
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
            4: [widgetGid, "c:wdg:aaa", "Widget", "Unit Price"],
            5: [widgetGid, "c:wdg:bbb", "Widget", "Unit  Price"],
          }),
          table: { endRowIndex: 6 },
        },
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: buildGridRows({
            0: ["c:wdg:aaa", "c:wdg:bbb"],
            3: ["Unit Price", "Unit  Price"],
            4: [1, 2],
          }),
          table: { endRowIndex: 5 },
        },
      ],
    });

    expect(() => initSyncedColumnConfigOperator().newColumnConfigs()).toThrow(
      /duplicate column name "unitPrice"/,
    );
  });
});

describe("ColumnConfigOperator.syncToSpreadsheet -> _updateProgrammaticValues", () => {
  const testSheetConfigRowWithApiAccess = [testSheetGid, "Item", true];

  function seedSheetConfigFixture() {
    return {
      sheetId: sheetConfigGid,
      title: "Sheet Config",
      rows: buildGridRows({
        0: sheetConfigColumnIdRow,
        3: sheetConfigHeaderRow,
        4: testSheetConfigRowWithApiAccess,
      }),
      table: { endRowIndex: 5, columnTypes: sheetConfigColumnTypes },
    };
  }

  it("corrects sheetTitle and header, emitting live samples", () => {
    stubSheetsService({
      sheets: [
        seedSheetConfigFixture(),
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
            4: [
              testSheetGid,
              "c:itm:corr01",
              "Stale Title",
              "Stale Header",
              true,
            ],
            5: [testSheetGid, "c:itm:corr02", "Item", "ID"],
          }),
          table: { endRowIndex: 6 },
        },
        {
          sheetId: testSheetGid,
          title: "Item",
          rows: buildGridRows({
            0: ["c:itm:corr01", "c:itm:corr02"],
            3: ["Amount", "ID"],
            4: [42, "xyz"],
          }),
          table: { endRowIndex: 5 },
        },
      ],
    });

    const operator = ColumnConfigOperator.init();
    syncColumnConfigOperator(operator);
    const identity = operator.sheet.columns("sheetTitle", "header");
    const emitted = operator.newColumnConfigs().item;

    expect(identity.sheetTitle.value(4)).toBe("Item");
    expect(identity.header.value(4)).toBe("Amount");
    expect(emitted?.amount).toMatchObject({
      valueName: "number",
      isFormula: false,
      emptyValueAllowed: true,
    });
    expect(operator.sheet.column("emptyValueAllowed").value(4)).toBe(true);

    expect(identity.sheetTitle.value(5)).toBe("Item");
    expect(identity.header.value(5)).toBe("ID");
    expect(emitted?.id).toMatchObject({
      valueName: "id",
      isFormula: false,
    });
  });

  it("fills in a row whose identity cells have never been filled in", () => {
    stubSheetsService({
      sheets: [
        seedSheetConfigFixture(),
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
            4: [testSheetGid, "c:itm:corr05", null, null],
          }),
          table: { endRowIndex: 5 },
        },
        {
          sheetId: testSheetGid,
          title: "Item",
          rows: buildGridRows({
            0: ["c:itm:corr05"],
            3: ["Amount"],
            4: [42],
          }),
          table: { endRowIndex: 5 },
        },
      ],
    });

    const operator = ColumnConfigOperator.init();
    syncColumnConfigOperator(operator);
    const identity = operator.sheet.columns("sheetTitle", "header");

    expect(identity.sheetTitle.value(4)).toBe("Item");
    expect(identity.header.value(4)).toBe("Amount");
    expect(operator.newColumnConfigs().item?.amount).toMatchObject({
      valueName: "number",
      isFormula: false,
    });
  });

  it("writes a floor column's ticked Empty value allowed back to FALSE and leaves a business column's tick", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: sheetConfigGid,
          title: "Sheet Config",
          rows: buildGridRows({
            0: sheetConfigColumnIdRow,
            3: sheetConfigHeaderRow,
            4: [testSheetGid, "Item", true],
            5: [sheetConfigGid, "Sheet Config", true],
          }),
          table: {
            endRowIndex: 6,
            endColumnIndex: 3,
            columnTypes: sheetConfigColumnTypes,
          },
        },
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
            4: [testSheetGid, "c:itm:corr01", "Item", "Amount", true],
            5: [
              sheetConfigGid,
              sc.sheetGid.columnId,
              "Sheet Config",
              sc.sheetGid.header,
              true,
            ],
          }),
          table: { endRowIndex: 6, columnTypes: columnConfigColumnTypes },
        },
        {
          sheetId: testSheetGid,
          title: "Item",
          rows: buildGridRows({
            0: ["c:itm:corr01"],
            3: ["Amount"],
            4: [42],
          }),
          table: { endRowIndex: 5 },
        },
      ],
    });

    const operator = ColumnConfigOperator.init();
    syncColumnConfigOperator(operator);
    const col = operator.sheet.columns("columnId", "emptyValueAllowed");
    const businessRow = operator.sheet.rowIndexesActiveWithData.find(
      (rowIndex) => col.columnId.value(rowIndex) === "c:itm:corr01",
    );
    const floorRow = operator.sheet.rowIndexesActiveWithData.find(
      (rowIndex) => col.columnId.value(rowIndex) === sc.sheetGid.columnId,
    );

    expect(businessRow).toBeDefined();
    expect(floorRow).toBeDefined();
    if (businessRow === undefined || floorRow === undefined) {
      throw new Error("Expected business and floor Column Config rows.");
    }
    expect(col.emptyValueAllowed.value(businessRow)).toBe(true);
    expect(col.emptyValueAllowed.value(floorRow)).toBe(false);
    expect(operator.declaredCellReport()).toContain(
      `Column Config · Empty value allowed · Sheet Config · ${sc.sheetGid.header} → FALSE`,
    );
  });

  it("detects a named valueConfig from the column's live data-validation formula", () => {
    stubSheetsService({
      sheets: [
        seedSheetConfigFixture(),
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
            4: [testSheetGid, "c:itm:corr03", "Item", "Description"],
          }),
          table: { endRowIndex: 5 },
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
    const identity = operator.sheet.columns("sheetTitle", "header");

    expect(valueTitles(operator, 1)).toEqual(["Transaction Description"]);
    expect(operator.newColumnConfigs().item?.description?.valueName).toBe(
      "transactionDescription",
    );
    expect(identity.sheetTitle.value(4)).toBe("Item");
    expect(identity.header.value(4)).toBe("Description");
  });

  it("detects a live formula and a date-formatted number", () => {
    stubSheetsService({
      sheets: [
        seedSheetConfigFixture(),
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
            4: [testSheetGid, "c:itm:corr04", "Item", "Due-by Date"],
          }),
          table: { endRowIndex: 5 },
        },
        {
          sheetId: testSheetGid,
          title: "Item",
          rows: buildGridRows({
            0: ["c:itm:corr04"],
            3: ["Due-by Date"],
            4: [{ value: 45000, isFormula: true, numberFormatType: "DATE" }],
          }),
          table: { endRowIndex: 5 },
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
    widgetGid,
    "c:wdg:aaa",
    "Widget",
    "Unit Price",
  ];
  const newSheetColumnConfigRow = [
    newSheetGid,
    "c:wdg:aaa",
    "Brand New Sheet",
    "Unit Price",
  ];

  function seedDuplicatedColumnIdOnTwoSheets(columnConfigRows: unknown[][]) {
    stubSheetsService({
      sheets: [
        {
          sheetId: sheetConfigGid,
          title: "Sheet Config",
          rows: buildGridRows({
            0: sheetConfigColumnIdRow,
            3: sheetConfigHeaderRow,
            4: [widgetGid, "Widget", true, ""],
            5: [newSheetGid, "Brand New Sheet", true, ""],
          }),
          table: { endRowIndex: 6, columnTypes: sheetConfigColumnTypes },
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
          table: { endRowIndex: 5 },
        },
        {
          sheetId: newSheetGid,
          title: "Brand New Sheet",
          rows: buildGridRows({
            0: ["c:wdg:aaa"],
            3: ["Unit Price"],
            4: [42],
          }),
          table: { endRowIndex: 5 },
        },
      ],
    });
  }

  function identitiesOnTheTwoSheets(operator: ColumnConfigOperator): string[] {
    const col = operator.sheet.columns("sheetGid", "columnId");
    return operator.sheet.rowIndexesActiveWithData
      .map((rowIndex) => [
        col.sheetGid.value(rowIndex),
        col.columnId.value(rowIndex),
      ])
      .filter(
        ([sheetGid]) => sheetGid === widgetGid || sheetGid === newSheetGid,
      )
      .map(([sheetGid, columnId]) => `${sheetGid}:${columnId}`);
  }

  it("appends a row for a column whose ID already has a row under another sheet", () => {
    seedDuplicatedColumnIdOnTwoSheets([propertyColumnConfigRow]);

    const operator = ColumnConfigOperator.init();
    syncColumnConfigOperator(operator);

    expect(identitiesOnTheTwoSheets(operator)).toEqual([
      `${widgetGid}:c:wdg:aaa`,
      `${newSheetGid}:c:wdg:aaa`,
    ]);
  });

  it("appends no row for a column whose identity already has one", () => {
    seedDuplicatedColumnIdOnTwoSheets([
      propertyColumnConfigRow,
      newSheetColumnConfigRow,
    ]);

    const operator = ColumnConfigOperator.init();
    syncColumnConfigOperator(operator);

    expect(identitiesOnTheTwoSheets(operator)).toEqual([
      `${widgetGid}:c:wdg:aaa`,
      `${newSheetGid}:c:wdg:aaa`,
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
    columnConfigRows[4 + index] = [testSheetGid, columnId, "Item", ""];
  });
  stubSheetsService({
    sheets: [
      {
        sheetId: sheetConfigGid,
        title: "Sheet Config",
        rows: buildGridRows({
          0: sheetConfigColumnIdRow,
          3: sheetConfigHeaderRow,
          4: [testSheetGid, "Item", true],
        }),
        table: { endRowIndex: 5, columnTypes: sheetConfigColumnTypes },
      },
      {
        sheetId: columnConfigGid,
        title: "Column Config",
        rows: buildGridRows(columnConfigRows),
        table: {
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
  const col = operator.sheet.columns("sheetGid", "columnId");
  const titles = operator.sheet.rowIndexesActiveWithData.flatMap((rowIndex) => {
    if (col.sheetGid.valueOrEmpty(rowIndex) !== testSheetGid) return [];
    return [
      operator.ss.raw
        .sheetMeta(testSheetGid)
        .columnByActiveId(col.columnId.value(rowIndex))
        .activeValueTitle(),
    ];
  });
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
      "1 column(s) across 1 sheet(s)",
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
      "1 column(s) across 1 sheet(s)",
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
      "1 column(s) across 1 sheet(s)",
    );
  });

  it("falls back to the sample for a column with no declared type, and counts it", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Amount"],
      topDataRow: [42],
    });

    expect(valueTitles(operator, 1)).toEqual(["number"]);
    expect(operator.untypedColumnsSummary()).toContain(
      "1 column(s) across 1 sheet(s)",
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
      "1 column(s) across 1 sheet(s)",
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
      "1 column(s) across 1 sheet(s)",
    );
  });

  it("keeps TRUE with only a number format guessed boolean and untyped", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Active"],
      topDataRow: [{ value: true, numberFormatType: "CURRENCY" }],
    });

    expect(valueTitles(operator, 1)).toEqual(["boolean"]);
    expect(operator.untypedColumnsSummary()).toContain(
      "1 column(s) across 1 sheet(s)",
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
      "1 column(s) across 1 sheet(s)",
    );
  });

  it("counts a formula column alongside the rest", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Balance"],
      topDataRow: [{ value: 42, isFormula: true }],
    });

    expect(operator.newColumnConfigs().item?.balance?.isFormula).toBe(true);
    expect(operator.untypedColumnsSummary()).toContain(
      "1 column(s) across 1 sheet(s)",
    );
  });

  it("summarises how many columns on how many sheets are still untyped", () => {
    const operator = syncColumnsUnderTest({
      headers: ["Amount", "Notes", "Moved In"],
      topDataRow: [42, "a note"],
      columnTypes: { 2: "DATE" },
    });

    expect(operator.untypedColumnsSummary()).toBe(
      "Succeeded, but 2 column(s) across 1 sheet(s) are untyped, so their " +
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
    const col = operator.sheet.columns("sheetTitle", "header");
    const emitted = operator.newColumnConfigs().item;

    expect(col.sheetTitle.value(4)).toBe("Item");
    expect(col.header.value(4)).toBe("Supplier Name");
    expect(emitted?.supplierName?.valueName).toBe("string");
    expect(col.header.value(5)).toBe("Amount");
    expect(emitted?.amount?.valueName).toBe("number");
  });

  it("records a column as no formula when the blank row proves nothing", () => {
    const operator = syncBlankSheetUnderTest({ headers: ["Amount"] });

    expect(operator.newColumnConfigs().item?.amount?.isFormula).toBe(false);
  });

  it("notes the sheets whose guesses had no sample row behind them", () => {
    const operator = syncBlankSheetUnderTest({ headers: ["Notes"] });

    expect(operator.untypedColumnsSummary()).toBe(
      "Succeeded, but 1 column(s) across 1 sheet(s) are untyped, so their " +
        "value names were guessed. See the execution log for the list. " +
        "On 1 of those sheet(s) the top data row was blank, so the guess had " +
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
    const col = operator.sheet.columns("header");

    expect(col.header.value(4)).toBe("Supplier Name");
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
          sheetId: sheetConfigGid,
          title: "Sheet Config",
          rows: buildGridRows({
            0: sheetConfigColumnIdRow,
            3: sheetConfigHeaderRow,
            // Api-access sheet: gets a missing column ID filled in.
            4: [testSheetGid, "Item", true, "tst"],

            5: [unresolvableGid, "Ghost", false, "gho"],
          }),
          table: { endRowIndex: 6 },
        },
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
          }),
          table: { endRowIndex: 5 },
        },
        {
          sheetId: testSheetGid,
          title: "Item",
          rows: buildGridRows({
            0: [""],
            3: ["Amount"],
            4: [42],
          }),
          table: { endRowIndex: 5 },
        },
      ],
    });

    const operator = ColumnConfigOperator.init();

    expect(() => syncColumnConfigOperator(operator)).not.toThrow();

    const colIdRow = operator.ss.raw.sheetMeta(testSheetGid).colIdRow;
    expect(colIdRow.valueOrEmpty(0)).not.toBe("");
  });

  it("adds a missing column ID when the columnId row has never had any grid data set", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: sheetConfigGid,
          title: "Sheet Config",
          rows: buildGridRows({
            0: sheetConfigColumnIdRow,
            3: sheetConfigHeaderRow,
            4: [testSheetGid, "Item", true, "tst"],
          }),
          table: { endRowIndex: 5 },
        },
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
          }),
          table: { endRowIndex: 5 },
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
          table: { endRowIndex: 5 },
        },
      ],
    });

    const operator = ColumnConfigOperator.init();

    expect(() => syncColumnConfigOperator(operator)).not.toThrow();

    const colIdRow = operator.ss.raw.sheetMeta(testSheetGid).colIdRow;
    expect(colIdRow.valueOrEmpty(0)).not.toBe("");
  });
});

describe("ColumnConfigOperator.syncToSpreadsheet -> _pruneColumnRows", () => {
  function seedColumnConfigDescribingItselfBelowABlankTopDataRow() {
    stubSheetsService({
      sheets: [
        {
          sheetId: sheetConfigGid,
          title: "Sheet Config",
          rows: buildGridRows({
            0: sheetConfigColumnIdRow,
            3: sheetConfigHeaderRow,
            4: [columnConfigGid, "Column Config", true, "ccf"],
          }),
          table: {
            endRowIndex: 5,
            endColumnIndex: 3,
            columnTypes: sheetConfigColumnTypes,
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
              columnConfigGid,
              cc.sheetGid.columnId,
              "Stale Title",
              "Stale Header",
            ],
          }),
          table: {
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
          sheetId: sheetConfigGid,
          title: "Sheet Config",
          rows: buildGridRows({
            0: sheetConfigColumnIdRow,
            3: sheetConfigHeaderRow,
            4: [columnConfigGid, "Column Config", false, "ccf"],
          }),
          table: {
            endRowIndex: 5,
            endColumnIndex: 3,
            columnTypes: sheetConfigColumnTypes,
          },
        },
        {
          sheetId: columnConfigGid,
          title: "Column Config",
          rows: buildGridRows({
            0: columnConfigColumnIdRow,
            3: columnConfigHeaderRow,
            4: [unresolvableGid, "c:???:eee", "Ghost", "Orphan Field"],
            5: [columnConfigGid, cc.sheetGid.columnId, "Column Config"],
          }),
          table: {
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
    expect(operator.sheet.column("columnId").hasValue("c:???:eee")).toBe(false);
    expect(operator.sheet.column("sheetGid").hasValue(columnConfigGid)).toBe(
      true,
    );
    expect(operator.newColumnConfigs().columnConfig).toBeDefined();
  });

  it("still resolves programmatic values for a sheet whose own top data row it pruned", () => {
    seedColumnConfigDescribingItselfBelowABlankTopDataRow();

    const operator = ColumnConfigOperator.init();
    syncColumnConfigOperator(operator);
    const col = operator.sheet.columns("sheetTitle", "header");
    const emitted = operator.newColumnConfigs().columnConfig;

    expect(operator.sheet.rowIndexesActive).not.toContain(4);
    expect(col.sheetTitle.value(5)).toBe("Column Config");
    expect(col.header.value(5)).toBe("Sheet GID");
    expect(emitted?.sheetGid).toMatchObject({
      valueName: "number",
      isFormula: false,
    });
  });
});
