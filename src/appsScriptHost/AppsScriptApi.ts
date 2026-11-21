import { AppsScript } from "../00_Source/GoogleSheets/AppsScript.ts";
import { GoogleSheetsAPI } from "../00_Source/GoogleSheets/GoogleSheetsAPI.ts";
import {
  hasInstalledRawSource,
  installRawSource,
} from "../00_Source/RawSource/installedRawSource.ts";
import type { DropdownStyleReset } from "../00_Source/RawSource/RawSource.ts";
import { Api, type AppSetup } from "../06_API/Api.ts";

// The Apps Script host's trigger glue, so an app's globals are one-liners. See docs/how-it-runs/apps-script-host.md.
export class AppsScriptApi {
  static handleSheetEdit(
    app: AppSetup,
    e: GoogleAppsScript.Events.SheetsOnEdit,
  ): void {
    Api.handleSheetEdit(app, AppsScript.sheetEdit(e), () =>
      installGoogleSheets(app.dropdownStyleReset),
    );
  }
}

function installGoogleSheets(
  dropdownStyleReset: DropdownStyleReset | undefined,
): void {
  if (!hasInstalledRawSource()) {
    installRawSource(GoogleSheetsAPI.forAppsScript(dropdownStyleReset));
  }
}
