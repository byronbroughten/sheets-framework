import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { type PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  type FakeSheetsService,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { fakeSpreadsheet } from "../testSupport/fakeSheetsService/fakeSpreadsheet";
import {
  itemColumns,
  itemGid,
  itemSheet,
  runItem,
} from "./syncCoordinatorTestSupport";

type SheetsRule = GoogleAppsScript.Sheets.Schema.ConditionalFormatRule;

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("SyncCoordinator.run conditional-format rule add", () => {
  const itemTable = Val.assert(itemSheet().tables?.[0], "the Item Table");
  const firstColIndex = Val.assert(itemTable.startColumnIndex, "first col");
  const headerRowIndex = Val.assert(itemTable.startRowIndex, "header row");
  const endRowIndex = Val.assert(itemTable.endRowIndex, "end row");
  const columnNames = Object.keys(itemColumns) as (keyof typeof itemColumns)[];
  const firstColumn = Val.assert(columnNames[0], "first column");

  function bodyRange(
    colOffset: number,
    width = 1,
  ): GoogleAppsScript.Sheets.Schema.GridRange {
    return {
      sheetId: itemGid,
      startRowIndex: headerRowIndex + 1,
      endRowIndex,
      startColumnIndex: firstColIndex + colOffset,
      endColumnIndex: firstColIndex + colOffset + width,
    };
  }

  function formulaRule(
    formula: string,
    red: number,
    range: GoogleAppsScript.Sheets.Schema.GridRange,
  ): SheetsRule {
    return {
      ranges: [range],
      booleanRule: {
        condition: {
          type: "CUSTOM_FORMULA",
          values: [{ userEnteredValue: formula }],
        },
        format: { backgroundColor: { red, green: 0, blue: 0 } },
      },
    };
  }

  const firstTrue = formulaRule("=TRUE", 1, bodyRange(0));
  const secondTrue = formulaRule("=TRUE", 1, bodyRange(1));
  const firstHandle = `${firstColumn} | =TRUE | background #ff0000`;

  function seedWithRules(): FakeSheetsService {
    return stubSheetsService({
      sheets: [{ ...itemSheet(), conditionalFormats: [firstTrue, secondTrue] }],
    });
  }

  function rulesLeft(service: FakeSheetsService): SheetsRule[] {
    return service.grid.sheet(itemGid).conditionalFormats ?? [];
  }

  const blackWhenFalse = {
    columns: [firstColumn],
    condition: { type: "CUSTOM_FORMULA" as const, values: ["=FALSE"] },
    format: { backgroundColor: "#000000" },
  };
  const blackWhenFalseRule = formulaRule("=FALSE", 0, bodyRange(0));

  const allOverThree = {
    columns: "all" as const,
    condition: { type: "NUMBER_GREATER" as const, values: ["3"] },
    format: { foregroundColor: "#0000ff" },
  };
  const allOverThreeRule: SheetsRule = {
    ranges: [bodyRange(0, columnNames.length)],
    booleanRule: {
      condition: {
        type: "NUMBER_GREATER",
        values: [{ userEnteredValue: "3" }],
      },
      format: {
        textFormat: { foregroundColor: { red: 0, green: 0, blue: 1 } },
      },
    },
  };

  function adding(
    add: (typeof blackWhenFalse | typeof allOverThree)[],
  ): PendingAlterations {
    return { tables: { item: { conditionalFormats: { add } } } };
  }

  it("sends two adds ahead of the rules already on their cells, the list's first first", () => {
    const service = seedWithRules();

    const { result, files } = runItem({
      isSending: true,
      alterations: adding([blackWhenFalse, allOverThree]),
    });

    expect(rulesLeft(service)).toEqual([
      blackWhenFalseRule,
      allOverThreeRule,
      firstTrue,
      secondTrue,
    ]);
    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      refusals: [],
      stops: [],
      plan: [
        {
          subject: { table: "item", path: "conditionalFormats.add[0]" },
          what: "add the rule",
        },
        {
          subject: { table: "item", path: "conditionalFormats.add[1]" },
          what: "add the rule",
        },
      ],
      didntLand: [],
    });
    const emitted = Val.assert(files, "emitted files");
    expect(emitted.conditionalFormatConfigs).toContain(
      `${firstColumn} | =FALSE | background #000000`,
    );
    expect(emitted.alterations).toContain("makePendingAlterations({})");
  });

  it("adds beside a replace on the same tab, the replacement landing on the rule it names", () => {
    const service = seedWithRules();

    const { result } = runItem({
      isSending: true,
      alterations: {
        tables: {
          item: {
            conditionalFormats: {
              add: [blackWhenFalse],
              replace: { [firstHandle]: allOverThree },
            },
          },
        },
      },
    });

    expect(rulesLeft(service)).toEqual([
      blackWhenFalseRule,
      allOverThreeRule,
      secondTrue,
    ]);
    expect(result).toMatchObject({ outcome: { kind: "sent" }, didntLand: [] });
  });

  it("adds a copy of a rule already there, landing as a second copy", () => {
    const service = seedWithRules();

    const { result } = runItem({
      isSending: true,
      alterations: adding([
        {
          ...blackWhenFalse,
          condition: { ...blackWhenFalse.condition, values: ["=TRUE"] },
          format: { backgroundColor: "#ff0000" },
        },
      ]),
    });

    expect(rulesLeft(service)).toEqual([firstTrue, firstTrue, secondTrue]);
    expect(result).toMatchObject({ outcome: { kind: "sent" }, didntLand: [] });
  });

  it("lands an add equal to a rule the same run replaces away", () => {
    const service = seedWithRules();

    const { result, files } = runItem({
      isSending: true,
      alterations: {
        tables: {
          item: {
            conditionalFormats: {
              add: [
                {
                  ...blackWhenFalse,
                  condition: { ...blackWhenFalse.condition, values: ["=TRUE"] },
                  format: { backgroundColor: "#ff0000" },
                },
              ],
              replace: { [firstHandle]: allOverThree },
            },
          },
        },
      },
    });

    expect(rulesLeft(service)).toEqual([
      firstTrue,
      allOverThreeRule,
      secondTrue,
    ]);
    expect(result).toMatchObject({ outcome: { kind: "sent" }, didntLand: [] });
    const emitted = Val.assert(files, "emitted files");
    expect(emitted.alterations).toContain("makePendingAlterations({})");
  });

  it("reports an add the refetch doesn't show under Didn't land, the entry staying pending", () => {
    const seeded = {
      ...itemSheet(),
      conditionalFormats: [firstTrue, secondTrue],
    };
    stubSheetsService({
      sheets: [seeded],
      afterEachBatchUpdate(spreadsheet) {
        spreadsheet.sheets = fakeSpreadsheet.init([seeded]).sheets;
      },
    });

    const { result, files } = runItem({
      isSending: true,
      alterations: adding([blackWhenFalse]),
    });

    const emitted = Val.assert(files, "emitted files");
    expect(emitted.alterations).toBeUndefined();
    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      landed: [],
      didntLand: [
        {
          subject: { table: "item", path: "conditionalFormats.add[0]" },
          what: "add the rule; Sheets didn't add it",
        },
      ],
    });
  });

  it("plans the adds on a dry run, sending nothing", () => {
    const service = seedWithRules();

    const { result } = runItem({
      isSending: false,
      alterations: adding([blackWhenFalse]),
    });

    expect(service.batchUpdateCount()).toBe(0);
    expect(rulesLeft(service)).toEqual([firstTrue, secondTrue]);
    expect(result.plan).toEqual([
      {
        subject: { table: "item", path: "conditionalFormats.add[0]" },
        what: "add the rule",
        effects: [],
      },
    ]);
  });

  it("adds a rule over a column the same run creates, in the one run that creates it", () => {
    const service = seedWithRules();

    const { result, files } = runItem({
      isSending: true,
      alterations: {
        tables: {
          item: {
            newColumns: [{ header: "Late fee" }],
            conditionalFormats: {
              add: [{ ...blackWhenFalse, columns: ["lateFee"] }],
            },
          },
        },
      },
    });

    expect(rulesLeft(service)).toEqual([
      formulaRule("=FALSE", 0, bodyRange(columnNames.length)),
      firstTrue,
      secondTrue,
    ]);
    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      refusals: [],
      didntLand: [],
    });
    const emitted = Val.assert(files, "emitted files");
    expect(emitted.conditionalFormatConfigs).toContain(
      "lateFee | =FALSE | background #000000",
    );
    expect(emitted.alterations).toContain("makePendingAlterations({})");
  });

  it("refuses before any read an add naming a column neither mapped nor created by the run", () => {
    const service = seedWithRules();

    const { result } = runItem({
      isSending: true,
      alterations: {
        tables: {
          item: {
            newColumns: [{ header: "Late fee" }],
            conditionalFormats: {
              add: [{ ...blackWhenFalse, columns: ["lateFees"] }],
            },
          },
        },
      },
    });

    expect(service.batchUpdateCount()).toBe(0);
    expect(result.refusals).toEqual([
      {
        kind: "unknownColumn",
        table: "item",
        path: "conditionalFormats.add[0].columns",
        column: "lateFees",
        columns: [...columnNames, "lateFee"],
      },
    ]);
  });

  it("refuses before any read an add naming a column two new headers both give", () => {
    const service = seedWithRules();

    const { result } = runItem({
      isSending: true,
      alterations: {
        tables: {
          item: {
            newColumns: [{ header: "Late fee" }, { header: "late Fee" }],
            conditionalFormats: {
              add: [{ ...blackWhenFalse, columns: ["lateFee"] }],
            },
          },
        },
      },
    });

    expect(service.batchUpdateCount()).toBe(0);
    expect(result.refusals).toEqual([
      {
        kind: "ambiguousColumn",
        table: "item",
        path: "conditionalFormats.add[0].columns",
        column: "lateFee",
        headers: ["Late fee", "late Fee"],
      },
    ]);
  });

  it("throws before any send on an add whose colour isn't six-digit hex", () => {
    const service = seedWithRules();

    expect(() =>
      runItem({
        isSending: true,
        alterations: adding([
          { ...blackWhenFalse, format: { backgroundColor: "black" } },
        ]),
      }),
    ).toThrow(/black/);
    expect(service.batchUpdateCount()).toBe(0);
  });
});
