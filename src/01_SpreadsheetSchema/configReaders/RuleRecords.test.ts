import { describe, expect, it } from "vitest";

import type {
  BooleanCondition,
  ConditionalFormat,
  ConditionalFormatRule,
} from "../../00_Source/RawSource/ConditionalFormat";
import type { GridRangeProps } from "../../00_Source/RawSource/RawSource";
import { SheetIndex } from "../../00_Source/RawSource/SheetIndex";
import { RuleRecords, type RuleRecordsTab } from "./RuleRecords";

const sheetGid = 7;
const red: ConditionalFormat = { backgroundColor: { red: 1 } };
const blueText: ConditionalFormat = { foregroundColor: { blue: 1 } };

// Header row 4 (index 3) over columns C:E, with body rows 5 to 10.
function itemTab(rules: ConditionalFormatRule[]): RuleRecordsTab {
  return {
    tabName: "items",
    sheetGid,
    tables: [
      {
        tableKey: "item",
        tableId: "table-1",
        layout: {
          headerTopLeft: {
            rowIndex: SheetIndex.row(3),
            colIndex: SheetIndex.col(2),
          },
          columnNamesLeftToRight: ["count", "newColumn", "price"],
        },
        endRowIndex: SheetIndex.row(10),
      },
    ],
    rules,
  };
}

function range(
  startRow: number,
  endRow: number,
  startCol: number,
  endCol: number,
): GridRangeProps {
  return {
    sheetId: sheetGid,
    startRowIndex: SheetIndex.row(startRow),
    endRowIndex: SheetIndex.row(endRow),
    startColumnIndex: SheetIndex.col(startCol),
    endColumnIndex: SheetIndex.col(endCol),
  };
}

function booleanRule(
  ranges: GridRangeProps[],
  condition: BooleanCondition,
  format: ConditionalFormat,
): ConditionalFormatRule {
  return { kind: "boolean", ranges, condition, format };
}

const newColumnBody = range(4, 10, 3, 4);
const overTwo: BooleanCondition = {
  type: "CUSTOM_FORMULA",
  formula: "=$D5>2",
};

describe("RuleRecords.build", () => {
  it("keys a Table's rule by its columns, brace-form condition and format, leaving out rows", () => {
    const records = RuleRecords.build([
      itemTab([
        booleanRule([newColumnBody], overTwo, red),
        booleanRule(
          [range(4, 10, 2, 5)],
          { type: "NUMBER_GREATER", value: 5 },
          blueText,
        ),
      ]),
    ]);
    expect(records).toEqual({
      tables: {
        item: {
          "newColumn | =${newColumn}>2 | background #ff0000": {
            kind: "boolean",
            tableId: "table-1",
            columns: ["newColumn"],
            condition: { type: "CUSTOM_FORMULA", formula: "=${newColumn}>2" },
            format: { backgroundColor: "#ff0000" },
          },
          "all | NUMBER_GREATER 5 | text #0000ff": {
            kind: "boolean",
            tableId: "table-1",
            columns: "all",
            condition: { type: "NUMBER_GREATER", value: 5 },
            format: { foregroundColor: "#0000ff" },
          },
        },
      },
      tabs: {},
    });
  });

  it("names the columns of a rule over several ranges left to right", () => {
    const records = RuleRecords.build([
      itemTab([
        booleanRule(
          [range(4, 10, 4, 5), range(4, 8, 2, 3)],
          { type: "TEXT_EQ", value: "x" },
          red,
        ),
      ]),
    ]);
    expect(Object.keys(records.tables.item ?? {})).toEqual([
      'count+price | TEXT_EQ "x" | background #ff0000',
    ]);
  });

  it("suffixes identical rules -2 and -3 in list order", () => {
    const rule = booleanRule([newColumnBody], overTwo, red);
    const records = RuleRecords.build([itemTab([rule, rule, rule])]);
    expect(Object.keys(records.tables.item ?? {})).toEqual([
      "newColumn | =${newColumn}>2 | background #ff0000",
      "newColumn | =${newColumn}>2 | background #ff0000-2",
      "newColumn | =${newColumn}>2 | background #ff0000-3",
    ]);
  });

  it("records a rule the brace form can't express as unmodelable, keyed by its raw content", () => {
    const records = RuleRecords.build([
      itemTab([
        {
          kind: "unmodelable",
          ranges: [range(4, 10, 2, 3)],
          index: 0,
          rawContent: '{"gradientRule":{}}',
        },
        booleanRule(
          [newColumnBody],
          { type: "CUSTOM_FORMULA", formula: "=$D6>2" },
          red,
        ),
        booleanRule([range(3, 10, 3, 4)], overTwo, red),
      ]),
    ]);
    const offRowContent =
      '{"condition":{"type":"CUSTOM_FORMULA","formula":"=$D6>2"},"format":{"backgroundColor":{"red":1}}}';
    const headerContent =
      '{"condition":{"type":"CUSTOM_FORMULA","formula":"=$D5>2"},"format":{"backgroundColor":{"red":1}}}';
    expect(records.tables.item).toEqual({
      'C5:C10 | unmodelable {"gradientRule":{}}': {
        kind: "unmodelable",
        tableId: "table-1",
        ranges: ["C5:C10"],
        rawContent: '{"gradientRule":{}}',
      },
      [`D5:D10 | unmodelable ${offRowContent}`]: {
        kind: "unmodelable",
        tableId: "table-1",
        ranges: ["D5:D10"],
        rawContent: offRowContent,
      },
      [`D4:D10 | unmodelable ${headerContent}`]: {
        kind: "unmodelable",
        tableId: "table-1",
        ranges: ["D4:D10"],
        rawContent: headerContent,
      },
    });
  });

  it("records a rule outside any Table under its tab, with its ranges and formula as written", () => {
    const records = RuleRecords.build([
      itemTab([booleanRule([range(0, 3, 0, 2)], overTwo, red)]),
    ]);
    expect(records).toEqual({
      tables: {},
      tabs: {
        items: {
          "A1:B3 | =$D5>2 | background #ff0000": {
            kind: "boolean",
            sheetGid,
            ranges: ["A1:B3"],
            condition: overTwo,
            format: { backgroundColor: "#ff0000" },
          },
        },
      },
    });
  });
});
