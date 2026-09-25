import { AppsScript } from "../00_Source/GoogleSheets/AppsScript";
import { GoogleSheetsAPI } from "../00_Source/GoogleSheets/GoogleSheetsAPI";
import {
  hasInstalledRawSource,
  installRawSource,
} from "../00_Source/RawSource/RawSource";
import type { FloorNotice } from "../05_Operators/ConfigSheetFloor/floorChangeNotice";
import { Api, type AppSetup } from "../06_API/Api";

const toastTimeoutSeconds = { floorNotice: 15, untilClosed: -1 } as const;

// The Apps Script host's trigger glue, so an app's globals are one-liners. See docs/how-it-runs.md.
export class AppsScriptApi {
  static handleSheetEdit(
    app: AppSetup,
    e: GoogleAppsScript.Events.SheetsOnEdit,
  ): void {
    Api.handleSheetEdit(app, AppsScript.sheetEdit(e), installGoogleSheets);
  }
  static handleSheetChange(
    app: AppSetup,
    e: GoogleAppsScript.Events.SheetsOnChange,
  ): void {
    const notice = Api.handleSheetChange(
      app,
      AppsScript.sheetChange(e.changeType),
      installGoogleSheets,
    );
    if (notice !== undefined) showFloorNotice(notice);
  }
}

function showFloorNotice({ title, message, untilClosed }: FloorNotice): void {
  AppsScript.toast(message, {
    title: `⚠️ ${title}`,
    timeoutSeconds: untilClosed
      ? toastTimeoutSeconds.untilClosed
      : toastTimeoutSeconds.floorNotice,
  });
}

function installGoogleSheets(): void {
  if (!hasInstalledRawSource()) {
    installRawSource(GoogleSheetsAPI.forAppsScript());
  }
}
