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

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("SyncCoordinator.run conditional-format rule removal", () => {
  const itemTable = Val.assert(itemSheet().tables?.[0], "the Item Table");
  const firstColIndex = Val.assert(itemTable.startColumnIndex, "first col");
  const headerRowIndex = Val.assert(itemTable.startRowIndex, "header row");
  const endRowIndex = Val.assert(itemTable.endRowIndex, "end row");
  const [firstColumn, secondColumn] = Object.keys(itemColumns);

  function bodyRange(
    colOffset: number,
  ): GoogleAppsScript.Sheets.Schema.GridRange {
    return {
      sheetId: itemGid,
      startRowIndex: headerRowIndex + 1,
      endRowIndex,
      startColumnIndex: firstColIndex + colOffset,
      endColumnIndex: firstColIndex + colOffset + 1,
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

  function removing(handles: string[]): PendingAlterations {
    return { tables: { item: { conditionalFormats: { remove: handles } } } };
  }

  function rulesLeft(service: FakeSheetsService): SheetsRule[] {
    return service.grid.sheet(itemGid).conditionalFormats ?? [];
  }

  it("removes the rule a handle names, leaving the Table's others", () => {
    const service = seedWithRules();

    const { result, files } = runItem({
      isSending: true,
      alterations: removing([firstHandle]),
    });

    expect(rulesLeft(service)).toEqual([gradientRule, trueRule(1)]);
    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      refusals: [],
      stops: [],
      didntLand: [],
    });
    const emitted = Val.assert(files, "emitted files");
    expect(emitted.conditionalFormatConfigs).not.toContain(firstHandle);
    expect(emitted.conditionalFormatConfigs).toContain(secondHandle);
    expect(emitted.alterations).toContain("makePendingAlterations({})");
  });

  it("removes an unmodelable rule by its handle, each delete indexed after the ones sent before it", () => {
    const service = seedWithRules();

    const { result } = runItem({
      isSending: true,
      alterations: removing([firstHandle, gradientHandle]),
    });

    expect(rulesLeft(service)).toEqual([trueRule(1)]);
    expect(result).toMatchObject({ outcome: { kind: "sent" }, stops: [] });
  });

  it("stops before any send on a handle no rule has now, listing the Table's handles and still refreshing the configs", () => {
    const service = seedWithRules();
    const changedHandle = `${firstColumn} | =FALSE | background #ff0000`;

    const { result, files } = runItem({
      isSending: true,
      alterations: removing([changedHandle]),
    });

    expect(service.batchUpdateCount()).toBe(0);
    expect(rulesLeft(service)).toHaveLength(3);
    expect(result).toMatchObject({
      outcome: { kind: "stopped" },
      refusals: [],
      stops: [
        {
          kind: "noSuchRule",
          table: "item",
          path: `conditionalFormats.remove["${changedHandle}"]`,
          handles: [firstHandle, gradientHandle, secondHandle],
        },
      ],
    });
    expect(files?.conditionalFormatConfigs).toContain(secondHandle);
  });
});
