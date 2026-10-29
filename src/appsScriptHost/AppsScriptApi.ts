import { AppsScript } from "../00_Source/GoogleSheets/AppsScript";
import { GoogleSheetsAPI } from "../00_Source/GoogleSheets/GoogleSheetsAPI";
import {
  hasInstalledRawSource,
  installRawSource,
} from "../00_Source/RawSource/RawSource";
import { Api, type AppSetup } from "../06_API/Api";

// The Apps Script host's trigger glue, so an app's globals are one-liners. See docs/how-it-runs/apps-script-host.md.
export class AppsScriptApi {
  static handleSheetEdit(
    app: AppSetup,
    e: GoogleAppsScript.Events.SheetsOnEdit,
  ): void {
    Api.handleSheetEdit(app, AppsScript.sheetEdit(e), installGoogleSheets);
  }
}

function installGoogleSheets(): void {
  if (!hasInstalledRawSource()) {
    installRawSource(GoogleSheetsAPI.forAppsScript());
  }
}
