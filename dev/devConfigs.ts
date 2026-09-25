import { columnConfigs } from "./generated/columnConfigs";
import { sheetConfigs } from "./generated/sheetConfigs";
import { spreadsheetConfig } from "./generated/spreadsheetConfig";
import { valueConfigs } from "./generated/valueConfigs";

export const devConfigs = {
  spreadsheetConfig,
  sheetConfigs,
  columnConfigs,
  valueConfigs,
};

declare module "../src/01_SpreadsheetSchema/configRegister" {
  interface Register {
    configs: typeof devConfigs;
  }
}
