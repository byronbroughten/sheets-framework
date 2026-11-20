import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

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
type ItemReplace = NonNullable<
  NonNullable<
    NonNullable<NonNullable<PendingAlterations["tables"]>["item"]>
  >["conditionalFormats"]
>["replace"];
type ItemRuleColumns = NonNullable<NonNullable<ItemReplace>[string]>["columns"];

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("SyncCoordinator.run conditional-format rule replacement", () => {
  const itemTable = Val.assert(itemSheet().tables?.[0], "the Item Table");
  const firstColIndex = Val.assert(itemTable.startColumnIndex, "first col");
  const headerRowIndex = Val.assert(itemTable.startRowIndex, "header row");
  const endRowIndex = Val.assert(itemTable.endRowIndex, "end row");
  const columnNames = Object.keys(itemColumns) as (keyof typeof itemColumns)[];
  const firstColumn = Val.assert(columnNames[0], "first column");
  const secondColumn = Val.assert(columnNames[1], "second column");

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

  function trueRule(colOffset: number): SheetsRule {
    return {
      ranges: [bodyRange(colOffset)],
      booleanRule: {
        condition: {
          type: "CUSTOM_FORMULA",
          values: [{ userEnteredValue: "=TRUE" }],
        },
        format: { backgroundColor: { red: 1, green: 0, blue: 0 } },
      },
    };
  }

  // A gradient has no brace form, so the configs record it as unmodelable.
  const gradientRule: SheetsRule = {
    ranges: [bodyRange(0)],
    gradientRule: {
      minpoint: { type: "MIN", color: { red: 1, green: 1, blue: 1 } },
      maxpoint: { type: "MAX", color: { red: 0, green: 1, blue: 0 } },
    },
  };

  const firstHandle = `${firstColumn} | =TRUE | background #ff0000`;
  const secondHandle = `${secondColumn} | =TRUE | background #ff0000`;
  const gradientHandle = `A5:A5 | unmodelable ${JSON.stringify({
    gradientRule: gradientRule.gradientRule,
  })}`;

  function seedWithRules(): FakeSheetsService {
    return stubSheetsService({
      sheets: [
        {
          ...itemSheet(),
          conditionalFormats: [trueRule(0), gradientRule, trueRule(1)],
        },
      ],
    });
  }

  function replacing(replace: ItemReplace): PendingAlterations {
    return { tables: { item: { conditionalFormats: { replace } } } };
  }

  function rulesLeft(service: FakeSheetsService): SheetsRule[] {
    return service.grid.sheet(itemGid).conditionalFormats ?? [];
  }

  const greenWhenFalse = {
    columns: [firstColumn],
    condition: { type: "CUSTOM_FORMULA" as const, values: ["=FALSE"] },
    format: { backgroundColor: "#00ff00" },
  };

  function greenWhenFalseRule(
    range: GoogleAppsScript.Sheets.Schema.GridRange,
  ): SheetsRule {
    return {
      ranges: [range],
      booleanRule: {
        condition: {
          type: "CUSTOM_FORMULA",
          values: [{ userEnteredValue: "=FALSE" }],
        },
        format: { backgroundColor: { red: 0, green: 1, blue: 0 } },
      },
    };
  }

  it("rewrites the rule a handle names at its index, leaving the Table's others in place", () => {
    const service = seedWithRules();

    const { result, files } = runItem({
      isSending: true,
      alterations: replacing({ [firstHandle]: greenWhenFalse }),
    });

    expect(rulesLeft(service)).toEqual([
      greenWhenFalseRule(bodyRange(0)),
      gradientRule,
      trueRule(1),
    ]);
    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      refusals: [],
      stops: [],
      didntLand: [],
    });
    const emitted = Val.assert(files, "emitted files");
    expect(emitted.conditionalFormatConfigs).not.toContain(firstHandle);
    expect(emitted.conditionalFormatConfigs).toContain(
      `${firstColumn} | =FALSE | background #00ff00`,
    );
    expect(emitted.alterations).toContain("makePendingAlterations({})");
  });

  it("spans every body column for `all`, compiling a comparison value", () => {
    const service = seedWithRules();

    const { result } = runItem({
      isSending: true,
      alterations: replacing({
        [secondHandle]: {
          columns: "all",
          condition: { type: "NUMBER_GREATER", values: ["3"] },
          format: { foregroundColor: "#0000ff" },
        },
      }),
    });

    expect(rulesLeft(service)[2]).toEqual({
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
    });
    expect(result).toMatchObject({ outcome: { kind: "sent" }, didntLand: [] });
  });

  it("replaces a rule beside a removal below it, at the index the removal leaves", () => {
    const service = seedWithRules();

    const { result } = runItem({
      isSending: true,
      alterations: {
        tables: {
          item: {
            conditionalFormats: {
              remove: [firstHandle],
              replace: { [secondHandle]: greenWhenFalse },
            },
          },
        },
      },
    });

    expect(rulesLeft(service)).toEqual([
      gradientRule,
      greenWhenFalseRule(bodyRange(0)),
    ]);
    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      stops: [],
      didntLand: [],
    });
  });

  it("stops before any send on a replace naming an unmodelable rule, which can only be removed", () => {
    const service = seedWithRules();

    const { result } = runItem({
      isSending: true,
      alterations: replacing({ [gradientHandle]: greenWhenFalse }),
    });

    expect(service.batchUpdateCount()).toBe(0);
    expect(rulesLeft(service)).toHaveLength(3);
    expect(result).toMatchObject({
      outcome: { kind: "stopped" },
      stops: [
        {
          kind: "unmodelableRule",
          table: "item",
          path: `conditionalFormats.replace["${gradientHandle}"]`,
          handle: gradientHandle,
        },
      ],
    });
  });

  it("stops before any send on a replace whose handle no rule has now", () => {
    const service = seedWithRules();
    const changedHandle = `${firstColumn} | =FALSE | background #ff0000`;

    const { result } = runItem({
      isSending: true,
      alterations: replacing({ [changedHandle]: greenWhenFalse }),
    });

    expect(service.batchUpdateCount()).toBe(0);
    expect(result).toMatchObject({
      outcome: { kind: "stopped" },
      stops: [
        {
          kind: "noSuchRule",
          table: "item",
          path: `conditionalFormats.replace["${changedHandle}"]`,
          handles: [firstHandle, gradientHandle, secondHandle],
        },
      ],
    });
  });

  it("refuses before any read a handle both removed and replaced", () => {
    const service = seedWithRules();

    const { result, files } = runItem({
      isSending: true,
      alterations: {
        tables: {
          item: {
            conditionalFormats: {
              remove: [firstHandle],
              replace: { [firstHandle]: greenWhenFalse },
            },
          },
        },
      },
    });

    expect(files).toBeUndefined();
    expect(service.batchUpdateCount()).toBe(0);
    expect(result).toMatchObject({
      outcome: { kind: "refused" },
      refusals: [
        {
          kind: "removedAndReplaced",
          table: "item",
          path: `conditionalFormats.replace["${firstHandle}"]`,
          removedAt: `conditionalFormats.remove["${firstHandle}"]`,
        },
      ],
    });
  });

  function renamingFirstAndReplacing(
    columns: ItemRuleColumns,
  ): PendingAlterations {
    return {
      tables: {
        item: {
          columns: { [firstColumn]: { header: "Renamed" } },
          conditionalFormats: {
            replace: { [secondHandle]: { ...greenWhenFalse, columns } },
          },
        },
      },
    };
  }

  it("refuses before any read a rule naming a column by the name the run's rename gives it", () => {
    const service = seedWithRules();

    const { result, files } = runItem({
      isSending: true,
      alterations: renamingFirstAndReplacing(["renamed"]),
    });

    expect(files).toBeUndefined();
    expect(service.batchUpdateCount()).toBe(0);
    expect(result.refusals).toEqual([
      {
        kind: "renamedColumn",
        table: "item",
        path: `conditionalFormats.replace["${secondHandle}"].columns`,
        column: "renamed",
        oldName: firstColumn,
      },
    ]);
  });

  it("compiles a rule naming a renamed column by its old name", () => {
    const service = seedWithRules();

    const { result } = runItem({
      isSending: true,
      alterations: renamingFirstAndReplacing([firstColumn]),
    });

    expect(rulesLeft(service)).toEqual([
      trueRule(0),
      gradientRule,
      greenWhenFalseRule(bodyRange(0)),
    ]);
    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      refusals: [],
      didntLand: [],
    });
  });

  it("refuses a bold format in replace and in add before any read", () => {
    const service = seedWithRules();
    const boldRule = {
      ...greenWhenFalse,
      format: { ...greenWhenFalse.format, bold: true, italic: true },
    };

    const { result, files } = runItem({
      isSending: true,
      alterations: {
        tables: {
          item: {
            conditionalFormats: {
              add: [boldRule],
              replace: { [firstHandle]: boldRule },
            },
          },
        },
      },
    });

    expect(files).toBeUndefined();
    expect(service.batchUpdateCount()).toBe(0);
    expect(result.refusals).toEqual([
      {
        kind: "notSendableYet",
        table: "item",
        path: "conditionalFormats.add[0].format.bold",
      },
      {
        kind: "notSendableYet",
        table: "item",
        path: "conditionalFormats.add[0].format.italic",
      },
      {
        kind: "notSendableYet",
        table: "item",
        path: `conditionalFormats.replace["${firstHandle}"].format.bold`,
      },
      {
        kind: "notSendableYet",
        table: "item",
        path: `conditionalFormats.replace["${firstHandle}"].format.italic`,
      },
    ]);
  });

  it("throws before any send on a colour that isn't six-digit hex", () => {
    const service = seedWithRules();

    expect(() =>
      runItem({
        isSending: true,
        alterations: replacing({
          [firstHandle]: {
            ...greenWhenFalse,
            format: { backgroundColor: "green" },
          },
        }),
      }),
    ).toThrow(/green/);
    expect(service.batchUpdateCount()).toBe(0);
  });
});
