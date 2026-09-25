import { Val } from "../../utils/Val";
import type { SheetChange } from "../PlatformEvents/sheetChange";
import type { SheetEdit } from "../PlatformEvents/sheetEdit";

interface ToastOptions {
  title: string;
  timeoutSeconds: number;
}

export class AppsScript {
  // A standalone script has no active spreadsheet; the framework's is always bound to its own.
  static boundSpreadsheetId(): string {
    return Val.assert(
      SpreadsheetApp.getActive(),
      "Active spreadsheet (bind the Apps Script project to its spreadsheet)",
    ).getId();
  }
  static sheetChange(
    changeType: GoogleAppsScript.Events.SheetsOnChange["changeType"],
  ): SheetChange | undefined {
    if (changeType === "REMOVE_GRID") return "sheetRemoved";
    if (changeType === "OTHER") return "other";
    return undefined;
  }
  // Google's event rows and columns are 1-based.
  static sheetEdit(e: GoogleAppsScript.Events.SheetsOnEdit): SheetEdit {
    return {
      sheetGid: e.range.getSheet().getSheetId(),
      rowIndexBase0: e.range.getRow() - 1,
      colIndexBase0: e.range.getColumn() - 1,
      value: e.value,
    };
  }
  static toast(message: string, { title, timeoutSeconds }: ToastOptions): void {
    SpreadsheetApp.getActive().toast(message, title, timeoutSeconds);
  }
  static get trigger(): {
    deleteAllTriggers(): void;
    addOnEdit(fnName: string): void;
    addOnChange(fnName: string): void;
    addFirstOfMonth: (fnName: string) => void;
    addEveryMinute: (fnName: string) => void;
  } {
    return {
      deleteAllTriggers(): void {
        ScriptApp.getProjectTriggers().forEach((trigger) => {
          ScriptApp.deleteTrigger(trigger);
        });
      },
      addOnEdit(fnName: string): void {
        ScriptApp.newTrigger(fnName)
          .forSpreadsheet(SpreadsheetApp.getActive())
          .onEdit()
          .create();
      },
      addOnChange(fnName: string): void {
        ScriptApp.newTrigger(fnName)
          .forSpreadsheet(SpreadsheetApp.getActive())
          .onChange()
          .create();
      },
      addFirstOfMonth: function (fnName: string) {
        ScriptApp.newTrigger(fnName).timeBased().onMonthDay(1).create();
      },
      addEveryMinute: function (fnName: string) {
        ScriptApp.newTrigger(fnName).timeBased().everyMinutes(1).create();
      },
    };
  }
}
