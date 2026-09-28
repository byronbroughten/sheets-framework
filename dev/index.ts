import { Api } from "../src/framework";
import { devEndpoints } from "./devEndpoints";
import { appConfigs } from "./generated/appConfigs";

const app = { configs: appConfigs, endpoints: devEndpoints };

// eslint-disable-next-line @typescript-eslint/no-unused-vars -- Apps Script calls it as a global trigger
function triggerOnEdit(e: GoogleAppsScript.Events.SheetsOnEdit): void {
  Api.handleSheetEdit(app, e);
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars -- Apps Script calls it as a global trigger
function triggerOnChange(e: GoogleAppsScript.Events.SheetsOnChange): void {
  Api.handleSheetChange(app, e);
}
