import { describe, expect, it } from "vitest";

import { stubScriptAndSpreadsheetApp } from "../../testSupport/fakeAppsScriptGlobals";
import { AppsScript } from "./AppsScript";

describe("AppsScript.boundSpreadsheetId", () => {
  it("reads the ID of the spreadsheet the script is bound to", () => {
    stubScriptAndSpreadsheetApp({ spreadsheetId: "abc123" });
    expect(AppsScript.boundSpreadsheetId()).toBe("abc123");
  });

  it("throws asking for a bound script when there is no active spreadsheet", () => {
    stubScriptAndSpreadsheetApp({ spreadsheetId: null });
    expect(() => AppsScript.boundSpreadsheetId()).toThrowError(
      "bind the Apps Script project to its spreadsheet",
    );
  });
});

describe("AppsScript.sheetChange", () => {
  it("decodes REMOVE_GRID as sheetRemoved", () => {
    expect(AppsScript.sheetChange("REMOVE_GRID")).toBe("sheetRemoved");
  });

  it("decodes OTHER as other", () => {
    expect(AppsScript.sheetChange("OTHER")).toBe("other");
  });

  it.each(["INSERT_ROW", "EDIT"] as const)(
    "ignores %s by returning undefined",
    (changeType) => {
      expect(AppsScript.sheetChange(changeType)).toBeUndefined();
    },
  );
});

describe("AppsScript.sheetEdit", () => {
  it("decodes an edit event into a 0-based SheetEdit", () => {
    const event = {
      value: "TRUE",
      range: {
        getRow: () => 5,
        getColumn: () => 3,
        getSheet: () => ({ getSheetId: () => 111 }),
      },
    } as unknown as GoogleAppsScript.Events.SheetsOnEdit;

    expect(AppsScript.sheetEdit(event)).toEqual({
      sheetGid: 111,
      rowIndexBase0: 4,
      colIndexBase0: 2,
      value: "TRUE",
    });
  });
});

describe("AppsScript.trigger", () => {
  it("addOnEdit schedules an onEdit trigger for the given function", () => {
    const { triggers } = stubScriptAndSpreadsheetApp();
    AppsScript.trigger.addOnEdit("triggerOnEdit");
    expect(triggers).toEqual([
      { handlerFunction: "triggerOnEdit", kind: "onEdit" },
    ]);
  });

  it("addOnChange schedules an onChange trigger for the given function, alongside the existing triggers", () => {
    const { triggers } = stubScriptAndSpreadsheetApp();
    AppsScript.trigger.addOnEdit("triggerOnEdit");
    AppsScript.trigger.addOnChange("triggerOnChange");
    expect(triggers).toEqual([
      { handlerFunction: "triggerOnEdit", kind: "onEdit" },
      { handlerFunction: "triggerOnChange", kind: "onChange" },
    ]);
  });

  it("addFirstOfMonth schedules a month-day-1 trigger", () => {
    const { triggers } = stubScriptAndSpreadsheetApp();
    AppsScript.trigger.addFirstOfMonth("monthlyJob");
    expect(triggers).toEqual([
      { handlerFunction: "monthlyJob", kind: "monthDay", detail: 1 },
    ]);
  });

  it("addEveryMinute schedules a 1-minute recurring trigger", () => {
    const { triggers } = stubScriptAndSpreadsheetApp();
    AppsScript.trigger.addEveryMinute("everyMinuteJob");
    expect(triggers).toEqual([
      { handlerFunction: "everyMinuteJob", kind: "everyMinutes", detail: 1 },
    ]);
  });

  it("deleteAllTriggers removes every currently scheduled trigger", () => {
    const { triggers } = stubScriptAndSpreadsheetApp();
    AppsScript.trigger.addOnEdit("a");
    AppsScript.trigger.addFirstOfMonth("b");
    expect(triggers).toHaveLength(2);

    AppsScript.trigger.deleteAllTriggers();
    expect(triggers).toHaveLength(0);
  });
});

describe("AppsScript.toast", () => {
  it("shows the message, title and timeout on the active spreadsheet", () => {
    const { toasts } = stubScriptAndSpreadsheetApp();
    AppsScript.toast("Press Undo now.", {
      title: "Value Config was deleted",
      timeoutSeconds: 15,
    });
    expect(toasts).toEqual([
      {
        message: "Press Undo now.",
        title: "Value Config was deleted",
        timeoutSeconds: 15,
      },
    ]);
  });
});
