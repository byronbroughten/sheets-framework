import { beforeEach, describe, expect, it, vi } from "vitest";

import { getColumnTraitByName } from "../01_SpreadsheetSchema/columnConfigsTypes";
import { installedConfigs } from "../01_SpreadsheetSchema/configRegister";
import { getSheetTraitByName } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import { ssConfigGet } from "../01_SpreadsheetSchema/spreadsheetConfigTypes";
import {
  stubLogger,
  stubScriptAndSpreadsheetApp,
} from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { AppsScriptApi } from "./AppsScriptApi";

const runItemGid = getSheetTraitByName("runItem", "sheetGid");
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

function onChangeEvent(
  changeType: GoogleAppsScript.Events.SheetsOnChange["changeType"],
): GoogleAppsScript.Events.SheetsOnChange {
  return { changeType } as GoogleAppsScript.Events.SheetsOnChange;
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
        ssConfigGet("actionRowIndexBase0"),
        resultColIndex,
        "TRUE",
      ),
    );
    expect(calls).toEqual(["result"]);
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
          ssConfigGet("actionRowIndexBase0"),
          resultColIndex,
          "TRUE",
        ),
      ),
    ).toThrow("bind the Apps Script project to its spreadsheet");
  });
});

describe("AppsScriptApi.handleSheetChange", () => {
  it("toasts a renamed Value Config's floor notice under a ⚠️ title, closing after 15 seconds", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: getSheetTraitByName("valueConfig", "sheetGid"),
          title: "Values",
        },
      ],
    });
    const { toasts } = stubScriptAndSpreadsheetApp();
    AppsScriptApi.handleSheetChange(
      { configs, endpoints: {} },
      onChangeEvent("OTHER"),
    );
    expect(toasts).toEqual([
      {
        message:
          'This tab keeps the name "Value Config". Your rename will switch back the next time configs sync.',
        title: "⚠️ Value Config is managed",
        timeoutSeconds: 15,
      },
    ]);
  });
  it("toasts a deleted Value Config's floor notice with no timeout, so it stays until closed", () => {
    stubSheetsService({ sheets: [] });
    const { toasts } = stubScriptAndSpreadsheetApp();
    AppsScriptApi.handleSheetChange(
      { configs, endpoints: {} },
      onChangeEvent("REMOVE_GRID"),
    );
    expect(toasts).toEqual([
      {
        message:
          "Press Undo (Ctrl+Z, or ⌘Z on a Mac) now to get it back with its data. If you don't, the next sync recreates it empty.",
        title: "⚠️ Value Config was deleted",
        timeoutSeconds: -1,
      },
    ]);
  });
  it("shows no toast for a change type the platform module doesn't name", () => {
    const { toasts } = stubScriptAndSpreadsheetApp();
    AppsScriptApi.handleSheetChange(
      { configs, endpoints: {} },
      onChangeEvent("EDIT"),
    );
    expect(toasts).toEqual([]);
  });
  it("reaches for the Google Sheets source when none is installed", async () => {
    vi.resetModules();
    const fresh = await import("./AppsScriptApi");
    stubScriptAndSpreadsheetApp({ spreadsheetId: null });
    expect(() =>
      fresh.AppsScriptApi.handleSheetChange(
        { configs, endpoints: {} },
        onChangeEvent("OTHER"),
      ),
    ).toThrow("bind the Apps Script project to its spreadsheet");
  });
});
