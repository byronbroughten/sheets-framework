import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SheetIndex } from "../00_Source/RawSource/SheetIndex";
import {
  BraceFormula,
  type RuleLayout,
} from "../01_SpreadsheetSchema/configReaders/BraceFormula";
import { type PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import {
  layoutBodyRows,
  layoutGid,
  layoutOrigins,
  layoutSheet,
} from "../02_SpreadsheetRaw/spreadsheetRawTestSupport";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  type FakeSheetsService,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import {
  itemColumns,
  itemGid,
  itemSheet,
  runItem,
  runLayout,
} from "./syncCoordinatorTestSupport";

type SheetsRule = GoogleAppsScript.Sheets.Schema.ConditionalFormatRule;
type GridRange = GoogleAppsScript.Sheets.Schema.GridRange;

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("SyncCoordinator.run conditional-format rule re-span", () => {
  const sheet = itemSheet({
    bodyRows: [{ id: "r:itm:one" }, { id: "r:itm:two" }, { id: "r:itm:three" }],
  });
  const itemTable = Val.assert(sheet.tables?.[0], "the Item Table");
  const firstColIndex = Val.assert(itemTable.startColumnIndex, "first col");
  const headerRowIndex = Val.assert(itemTable.startRowIndex, "header row");
  const endRowIndex = Val.assert(itemTable.endRowIndex, "end row");
  // Two body rows short of the Table's three, as rows typed below it leave a rule.
  const shortEndRowIndex = endRowIndex - 1;
  const columnNames = Object.keys(itemColumns) as (keyof typeof itemColumns)[];
  const firstColumn = Val.assert(columnNames[0], "first column");
  const secondColumn = Val.assert(columnNames[1], "second column");
  const duplicateFormula = `=COUNTIF(\${${firstColumn}},\${@${firstColumn}})>1`;

  function layoutEndingAt(layoutEndRowIndex: number): RuleLayout {
    return {
      headerTopLeft: {
        rowIndex: SheetIndex.row(headerRowIndex),
        colIndex: SheetIndex.col(firstColIndex),
      },
      endRowIndex: SheetIndex.row(layoutEndRowIndex),
      columnNamesLeftToRight: columnNames,
    };
  }

  function bodyRange(colOffset: number, rangeEndRowIndex: number): GridRange {
    return {
      sheetId: itemGid,
      startRowIndex: headerRowIndex + 1,
      endRowIndex: rangeEndRowIndex,
      startColumnIndex: firstColIndex + colOffset,
      endColumnIndex: firstColIndex + colOffset + 1,
    };
  }

  function formulaRule(range: GridRange, formula: string): SheetsRule {
    return {
      ranges: [range],
      booleanRule: {
        condition: {
          type: "CUSTOM_FORMULA",
          values: [{ userEnteredValue: formula }],
        },
        format: { backgroundColor: { red: 1, green: 0, blue: 0 } },
      },
    };
  }

  function duplicateRuleEndingAt(ruleEndRowIndex: number): SheetsRule {
    return formulaRule(
      bodyRange(0, ruleEndRowIndex),
      BraceFormula.compileForRule({
        tableKey: "item",
        formula: duplicateFormula,
        layout: layoutEndingAt(ruleEndRowIndex),
      }),
    );
  }

  const trueRule = formulaRule(bodyRange(1, endRowIndex), "=TRUE");
  const duplicateHandle = `${firstColumn} | ${duplicateFormula} | background #ff0000`;
  const trueHandle = `${secondColumn} | =TRUE | background #ff0000`;

  function seedWith(rules: SheetsRule[]): FakeSheetsService {
    return stubSheetsService({
      sheets: [{ ...sheet, conditionalFormats: rules }],
    });
  }

  function rulesLeft(service: FakeSheetsService): SheetsRule[] {
    return service.grid.sheet(itemGid).conditionalFormats ?? [];
  }

  function runWith(alterations: PendingAlterations = {}) {
    return runItem({ isSending: true, alterations });
  }

  it("stretches a short rule to the last body row, recompiling its formula, as the run's only send", () => {
    const service = seedWith([
      duplicateRuleEndingAt(shortEndRowIndex),
      trueRule,
    ]);

    const { result } = runWith();

    expect(rulesLeft(service)).toEqual([
      duplicateRuleEndingAt(endRowIndex),
      trueRule,
    ]);
    expect(service.batchUpdateCount()).toBe(1);
    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      stops: [],
      didntLand: [],
    });
    expect(result.plan).toEqual([
      {
        subject: {
          table: "item",
          path: `conditionalFormats["${duplicateHandle}"]`,
        },
        what: "re-span the rule",
        effects: [
          {
            kind: "ruleRespanned",
            handle: duplicateHandle,
            fromRow: shortEndRowIndex,
            toRow: endRowIndex,
          },
        ],
      },
    ]);
  });

  it("re-spans at the index a removal above it and an add leave, in the same batch", () => {
    const service = seedWith([
      trueRule,
      duplicateRuleEndingAt(shortEndRowIndex),
    ]);
    const greenWhenFalse = {
      columns: [secondColumn],
      condition: { type: "CUSTOM_FORMULA" as const, values: ["=FALSE"] },
      format: { backgroundColor: "#00ff00" },
    };

    const { result } = runWith({
      tables: {
        item: {
          conditionalFormats: { remove: [trueHandle], add: [greenWhenFalse] },
        },
      },
    });

    expect(rulesLeft(service)).toEqual([
      {
        ranges: [bodyRange(1, endRowIndex)],
        booleanRule: {
          condition: {
            type: "CUSTOM_FORMULA",
            values: [{ userEnteredValue: "=FALSE" }],
          },
          format: { backgroundColor: { red: 0, green: 1, blue: 0 } },
        },
      },
      duplicateRuleEndingAt(endRowIndex),
    ]);
    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      stops: [],
      didntLand: [],
    });
  });

  it("leaves a short rule with no brace form as it is", () => {
    const outsideRule = formulaRule(bodyRange(0, shortEndRowIndex), "=$Z$1>0");
    const service = seedWith([outsideRule]);

    const { result } = runWith();

    expect(rulesLeft(service)).toEqual([outsideRule]);
    expect(service.batchUpdateCount()).toBe(0);
    expect(result.plan).toEqual([]);
  });

  function numberRule(
    type: string,
    range: GridRange,
    userEnteredValue: string,
  ): SheetsRule {
    return {
      ranges: [range],
      booleanRule: {
        condition: { type, values: [{ userEnteredValue }] },
        format: { backgroundColor: { red: 1, green: 0, blue: 0 } },
      },
    };
  }

  it.each([
    ["NUMBER_EQ", "FALSE"],
    ["NUMBER_EQ", "Yes"],
    ["NUMBER_NOT_EQ", "Yes"],
    ["NUMBER_GREATER", "FALSE"],
  ])(
    "leaves a short %s rule holding %s as it is, since Google refuses to rewrite it",
    (type, userEnteredValue) => {
      const textRule = numberRule(
        type,
        bodyRange(1, shortEndRowIndex),
        userEnteredValue,
      );
      const service = seedWith([textRule]);

      const { result } = runWith();

      expect(rulesLeft(service)).toEqual([textRule]);
      expect(service.batchUpdateCount()).toBe(0);
      expect(result.plan).toEqual([]);
    },
  );

  it.each(["0", "1.50"])(
    "still re-spans a short NUMBER_EQ rule holding %s",
    (userEnteredValue) => {
      const service = seedWith([
        numberRule(
          "NUMBER_EQ",
          bodyRange(1, shortEndRowIndex),
          userEnteredValue,
        ),
      ]);

      const { result } = runWith();

      expect(rulesLeft(service)).toEqual([
        numberRule("NUMBER_EQ", bodyRange(1, endRowIndex), userEnteredValue),
      ]);
      expect(result).toMatchObject({
        outcome: { kind: "sent" },
        didntLand: [],
      });
    },
  );

  it("stops partway when Google refuses the rule flush, the checkbox's flush before it landed, and still returns the files", () => {
    const googleRefusal =
      "Invalid requests[0].updateConditionalFormatRule: Invalid ConditionValue.userEnteredValue: FALSE for ConditionType: NUMBER_EQ";
    let batchCount = 0;
    const service = stubSheetsService({
      sheets: [
        {
          ...itemSheet({
            bodyRows: [
              { id: "r:itm:one" },
              { id: "r:itm:two" },
              { id: "r:itm:three" },
            ],
            actionRowCells: {},
          }),
          conditionalFormats: [duplicateRuleEndingAt(shortEndRowIndex)],
        },
      ],
      beforeEachBatchUpdate() {
        batchCount++;
        if (batchCount === 2) throw new Error(googleRefusal);
      },
    });

    const { result, files } = runWith();

    expect(files).toBeDefined();
    expect(rulesLeft(service)).toEqual([
      duplicateRuleEndingAt(shortEndRowIndex),
    ]);
    expect(result.outcome).toEqual({
      kind: "stoppedPartway",
      landedFlushes: 1,
      googleRefusal,
    });
  });

  it("records a column deletion in the refused rule flush's batch as not landed, keeping the column in the configs", () => {
    const service = stubSheetsService({
      sheets: [
        {
          ...sheet,
          conditionalFormats: [duplicateRuleEndingAt(shortEndRowIndex)],
        },
      ],
      beforeEachBatchUpdate() {
        throw new Error(
          "Invalid requests[1].updateConditionalFormatRule: Invalid ConditionValue.userEnteredValue: FALSE for ConditionType: NUMBER_EQ",
        );
      },
    });

    const { result, files } = runWith({
      tables: { item: { columns: { optionalNote: { deleted: true } } } },
    });

    const headerRow =
      service.grid.sheet(itemGid).values()[headerRowIndex] ?? [];
    expect(headerRow).toContain(itemColumns.optionalNote.header);
    expect(files?.columnConfigs).toContain(itemColumns.optionalNote.columnId);
    expect(result).toMatchObject({
      outcome: { kind: "stoppedBeforeSending" },
      didntLand: [{ subject: { table: "item", path: "deleted" } }],
    });
  });

  it("leaves a short rule a replace names to the replace", () => {
    const service = seedWith([duplicateRuleEndingAt(shortEndRowIndex)]);

    const { result } = runWith({
      tables: {
        item: {
          conditionalFormats: {
            replace: {
              [duplicateHandle]: {
                columns: [firstColumn],
                condition: { type: "CUSTOM_FORMULA", values: ["=FALSE"] },
                format: { backgroundColor: "#00ff00" },
              },
            },
          },
        },
      },
    });

    expect(rulesLeft(service)).toEqual([
      {
        ranges: [bodyRange(0, endRowIndex)],
        booleanRule: {
          condition: {
            type: "CUSTOM_FORMULA",
            values: [{ userEnteredValue: "=FALSE" }],
          },
          format: { backgroundColor: { red: 0, green: 1, blue: 0 } },
        },
      },
    ]);
    expect(
      result.plan.flatMap(({ effects }) =>
        effects.filter(({ kind }) => kind === "ruleRespanned"),
      ),
    ).toEqual([]);
  });
});

describe("SyncCoordinator.run rule re-span beside a column insert on the same tab", () => {
  // No column references: a compiled reference is $-anchored, whose shift under an insert the fake doesn't replay in a moved rule.
  const amountFormula = "=TRUE";
  const { layoutLeft, layoutRight } = layoutOrigins;
  const headerRowIndex = layoutLeft.startRowIndex;
  const leftEndRowIndex = headerRowIndex + 1 + layoutBodyRows.layoutLeft.length;
  const rightEndRowIndex =
    headerRowIndex + 1 + layoutBodyRows.layoutRight.length;

  interface AmountRuleProps {
    tableKey: "layoutLeft" | "layoutRight";
    startColIndex: number;
    columnNamesLeftToRight: string[];
    ruleEndRowIndex: number;
  }

  // The amount rule as a layout compiles it: its range and formula both follow the Table's start column and end row.
  function amountRule({
    tableKey,
    startColIndex,
    columnNamesLeftToRight,
    ruleEndRowIndex,
  }: AmountRuleProps): SheetsRule {
    const amountColIndex =
      startColIndex + columnNamesLeftToRight.indexOf("amount");
    return {
      ranges: [
        {
          sheetId: layoutGid,
          startRowIndex: headerRowIndex + 1,
          endRowIndex: ruleEndRowIndex,
          startColumnIndex: amountColIndex,
          endColumnIndex: amountColIndex + 1,
        },
      ],
      booleanRule: {
        condition: {
          type: "CUSTOM_FORMULA",
          values: [
            {
              userEnteredValue: BraceFormula.compileForRule({
                tableKey,
                formula: amountFormula,
                layout: {
                  headerTopLeft: {
                    rowIndex: SheetIndex.row(headerRowIndex),
                    colIndex: SheetIndex.col(startColIndex),
                  },
                  endRowIndex: SheetIndex.row(ruleEndRowIndex),
                  columnNamesLeftToRight,
                },
              }),
            },
          ],
        },
        format: { backgroundColor: { red: 1, green: 0, blue: 0 } },
      },
    };
  }

  it("holds back a short rule on the Table the insert shifts, leaving it where Sheets moved it, and rewrites the inserted Table's short rule to its last body row", () => {
    const shortLeft = amountRule({
      tableKey: "layoutLeft",
      startColIndex: layoutLeft.startColumnIndex,
      columnNamesLeftToRight: ["entry", "amount"],
      ruleEndRowIndex: leftEndRowIndex - 1,
    });
    const shortRight = amountRule({
      tableKey: "layoutRight",
      startColIndex: layoutRight.startColumnIndex,
      columnNamesLeftToRight: ["entry", "amount"],
      ruleEndRowIndex: rightEndRowIndex - 1,
    });
    const service = stubSheetsService({
      sheets: [
        { ...layoutSheet(), conditionalFormats: [shortLeft, shortRight] },
      ],
    });

    const { result } = runLayout({
      tables: {
        layoutLeft: {
          newColumns: [{ columns: [{ header: "Note" }], before: "amount" }],
        },
      },
    });

    expect(
      service.grid
        .sheet(layoutGid)
        .tables.find(({ tableId }) => tableId === "layoutRight")?.range,
    ).toMatchObject({ startColumnIndex: layoutRight.startColumnIndex + 1 });
    expect(service.grid.sheet(layoutGid).conditionalFormats).toEqual([
      amountRule({
        tableKey: "layoutLeft",
        startColIndex: layoutLeft.startColumnIndex,
        columnNamesLeftToRight: ["entry", "note", "amount"],
        ruleEndRowIndex: leftEndRowIndex,
      }),
      amountRule({
        tableKey: "layoutRight",
        startColIndex: layoutRight.startColumnIndex + 1,
        columnNamesLeftToRight: ["entry", "amount"],
        ruleEndRowIndex: rightEndRowIndex - 1,
      }),
    ]);
    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      stops: [],
      didntLand: [],
    });
  });
});
