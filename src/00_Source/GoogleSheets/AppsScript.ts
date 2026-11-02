import { Val } from "@byronbroughten/utils/val";

import type { SheetEdit } from "../PlatformEvents/sheetEdit";
import { SheetIndex } from "../RawSource/SheetIndex";

export class AppsScript {
  // A standalone script has no active spreadsheet; the framework's is always bound to its own.
  static boundSpreadsheetId(): string {
    return Val.assert(
      SpreadsheetApp.getActive(),
      "Active spreadsheet (bind the Apps Script project to its spreadsheet)",
    ).getId();
  }
  // The code assumes the event range's rows and columns are 1-based; documented for columns: docs/sheets-behaviour.md#edit-event-ranges.
  static sheetEdit(e: GoogleAppsScript.Events.SheetsOnEdit): SheetEdit {
    return {
      sheetGid: e.range.getSheet().getSheetId(),
      rowIndexBase0: SheetIndex.row(e.range.getRow() - 1),
      colIndexBase0: SheetIndex.col(e.range.getColumn() - 1),
      value: e.value,
    };
  }
  static get trigger(): {
    deleteAllTriggers(): void;
    addOnEdit(fnName: string): void;
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
      addFirstOfMonth: function (fnName: string) {
        ScriptApp.newTrigger(fnName).timeBased().onMonthDay(1).create();
      },
      addEveryMinute: function (fnName: string) {
        ScriptApp.newTrigger(fnName).timeBased().everyMinutes(1).create();
      },
    };
  }
}
