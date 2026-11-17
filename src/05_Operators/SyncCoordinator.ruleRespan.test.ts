import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SheetIndex } from "../00_Source/RawSource/SheetIndex";
import {
  BraceFormula,
  type RuleLayout,
} from "../01_SpreadsheetSchema/configReaders/BraceFormula";
import { type PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
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
