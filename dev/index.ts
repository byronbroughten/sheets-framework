import { Api } from "../src/framework.ts";
import { devEndpoints } from "./devEndpoints.ts";
import { appConfigs } from "./generated/appConfigs.ts";

const app = { configs: appConfigs, endpoints: devEndpoints };

// eslint-disable-next-line @typescript-eslint/no-unused-vars -- Apps Script calls it as a global trigger
function triggerOnEdit(e: GoogleAppsScript.Events.SheetsOnEdit): void {
  Api.handleSheetEdit(app, e);
}
