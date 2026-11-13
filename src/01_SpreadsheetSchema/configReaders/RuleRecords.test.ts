import { describe, expect, it } from "vitest";

import type {
  BooleanCondition,
  ConditionalFormat,
  ConditionalFormatRule,
  ModelableConditionalFormatRule,
} from "../../00_Source/RawSource/ConditionalFormat";
import type { GridRangeProps } from "../../00_Source/RawSource/RawSource";
import { SheetIndex } from "../../00_Source/RawSource/SheetIndex";
import { RuleRecords, type RuleRecordsTab } from "./RuleRecords";

const sheetGid = 7;
const red: ConditionalFormat = { backgroundColor: { red: 1 } };
const blueText: ConditionalFormat = { foregroundColor: { blue: 1 } };

// Header row 4 (index 3) over columns C:E, with body rows 5 to 10.
function itemTab(
  rules: ConditionalFormatRule[],
  columnNamesLeftToRight: (string | undefined)[] = [
    "count",
    "newColumn",
    "price",
  ],
): RuleRecordsTab {
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
          columnNamesLeftToRight,
        },
        endRowIndex: SheetIndex.row(10),
      },
    ],
    rules,
  };
}

// Each pair is a start and an exclusive end index.
interface RangeBounds {
  rows: [number, number];
  cols: [number, number];
}

function range({ rows, cols }: RangeBounds): GridRangeProps {
  return {
    sheetId: sheetGid,
    startRowIndex: SheetIndex.row(rows[0]),
    endRowIndex: SheetIndex.row(rows[1]),
    startColumnIndex: SheetIndex.col(cols[0]),
    endColumnIndex: SheetIndex.col(cols[1]),
  };
}

function booleanRule(
  rule: Omit<ModelableConditionalFormatRule, "kind">,
): ConditionalFormatRule {
  return { kind: "boolean", ...rule };
}

const newColumnBody = range({ rows: [4, 10], cols: [3, 4] });
const overTwo: BooleanCondition = {
  type: "CUSTOM_FORMULA",
  formula: "=$D5>2",
};

describe("RuleRecords.build", () => {
  it("keys a Table's rule by its columns, brace-form condition and format, leaving out rows", () => {
    const records = RuleRecords.build([
      itemTab([
        booleanRule({
          ranges: [newColumnBody],
          condition: overTwo,
          format: red,
        }),
        booleanRule({
          ranges: [range({ rows: [4, 10], cols: [2, 5] })],
          condition: { type: "NUMBER_GREATER", value: 5 },
          format: blueText,
        }),
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

  it("records a rule over a column with no name this run as unmodelable, and still names a rule clear of it", () => {
    const countBody = range({ rows: [4, 10], cols: [2, 3] });
    const records = RuleRecords.build([
      itemTab(
        [
          booleanRule({
            ranges: [newColumnBody],
            condition: overTwo,
            format: red,
          }),
          booleanRule({
            ranges: [countBody],
            condition: { type: "NUMBER_GREATER", value: 5 },
            format: blueText,
          }),
        ],
        ["count", undefined, "price"],
      ),
    ]);
    expect(Object.values(records.tables.item ?? {})).toMatchObject([
      { kind: "unmodelable", tableId: "table-1", ranges: ["D5:D10"] },
      { kind: "boolean", tableId: "table-1", columns: ["count"] },
    ]);
  });

  it("names the columns of a rule over several ranges left to right", () => {
    const records = RuleRecords.build([
      itemTab([
        booleanRule({
          ranges: [
            range({ rows: [4, 10], cols: [4, 5] }),
            range({ rows: [4, 8], cols: [2, 3] }),
          ],
          condition: { type: "TEXT_EQ", value: "x" },
          format: red,
        }),
      ]),
    ]);
    expect(Object.keys(records.tables.item ?? {})).toEqual([
      'count+price | TEXT_EQ "x" | background #ff0000',
    ]);
  });

  it("suffixes identical rules -2 and -3 in list order", () => {
    const rule = booleanRule({
      ranges: [newColumnBody],
      condition: overTwo,
      format: red,
    });
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
          ranges: [range({ rows: [4, 10], cols: [2, 3] })],
          index: 0,
          rawContent: '{"gradientRule":{}}',
        },
        booleanRule({
          ranges: [newColumnBody],
          condition: { type: "CUSTOM_FORMULA", formula: "=$D6>2" },
          format: red,
        }),
        booleanRule({
          ranges: [range({ rows: [3, 10], cols: [3, 4] })],
          condition: overTwo,
          format: red,
        }),
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
      itemTab([
        booleanRule({
          ranges: [range({ rows: [0, 3], cols: [0, 2] })],
          condition: overTwo,
          format: red,
        }),
      ]),
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
