import { beforeEach, describe, expect, it, vi } from "vitest";

import { getColumnTraitByName } from "../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import { installedConfigs } from "../01_SpreadsheetSchema/configReaders/configRegister";
import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { expectedSheetLayout } from "../testSupport/expectedSheetLayout";
import {
  stubLogger,
  stubScriptAndSpreadsheetApp,
} from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { AppsScriptApi } from "./AppsScriptApi";

const runItemGid = getTableTraitByName("runItem", "sheetGid");
const resultColIndex = 1;
const configs = installedConfigs();

// Google's event rows and columns are 1-based.
function onEditEvent(
  sheetGid: number,
  rowIndexBase0: number,
  colIndexBase0: number,
  value: string,
): GoogleAppsScript.Events.SheetsOnEdit {
  return {
    range: {
      getSheet: () => ({ getSheetId: () => sheetGid }),
      getRow: () => rowIndexBase0 + 1,
      getColumn: () => colIndexBase0 + 1,
    },
    value,
  } as unknown as GoogleAppsScript.Events.SheetsOnEdit;
}

beforeEach(() => {
  stubLogger();
});

describe("AppsScriptApi.handleSheetEdit", () => {
  it("decodes the event and dispatches an action-row tick through the run", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: runItemGid,
          title: "Run item",
          rows: buildGridRows({
            0: [
              getColumnTraitByName("runItem", "id", "columnId"),
              getColumnTraitByName("runItem", "result", "columnId"),
            ],
            3: ["ID", "Result"],
            4: ["r:rit:row4", ""],
          }),
          table: { endRowIndex: 5 },
        },
      ],
    });
    const calls: string[] = [];
    AppsScriptApi.handleSheetEdit(
      {
        configs,
        endpoints: {
          runItem_result: {
            action: () => {
              calls.push("result");
            },
          },
        },
      },
      onEditEvent(
        runItemGid,
        expectedSheetLayout.actionRowIndex,
        resultColIndex,
        "TRUE",
      ),
    );
    expect(calls).toEqual(["result"]);
  });

  it("decodes a data-row tick as no match, costing no fetch", () => {
    const { getByDataFilterCalls } = stubSheetsService({
      sheets: [
        {
          sheetId: runItemGid,
          title: "Run item",
          rows: buildGridRows({
            3: ["ID", "Result"],
            4: ["r:rit:row4", true],
          }),
          table: { endRowIndex: 5 },
        },
      ],
    });
    const calls: string[] = [];
    AppsScriptApi.handleSheetEdit(
      {
        configs,
        endpoints: {
          runItem_result: {
            action: () => {
              calls.push("result");
            },
          },
        },
      },
      onEditEvent(
        runItemGid,
        expectedSheetLayout.topDataRowIndex,
        resultColIndex,
        "TRUE",
      ),
    );
    expect(calls).toEqual([]);
    expect(getByDataFilterCalls).toHaveLength(0);
  });
});

describe("AppsScriptApi.handleSheetEdit, with no source installed", () => {
  it("reaches for the Google Sheets source on an action-row tick", async () => {
    vi.resetModules();
    const fresh = await import("./AppsScriptApi");
    stubScriptAndSpreadsheetApp({ spreadsheetId: null });
    expect(() =>
      fresh.AppsScriptApi.handleSheetEdit(
        { configs, endpoints: {} },
        onEditEvent(
          runItemGid,
          expectedSheetLayout.actionRowIndex,
          resultColIndex,
          "TRUE",
        ),
      ),
    ).toThrow("bind the Apps Script project to its spreadsheet");
  });
});
