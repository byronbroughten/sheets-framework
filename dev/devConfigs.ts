import { columnConfigs } from "./generated/columnConfigs";
import { sheetConfigs } from "./generated/sheetConfigs";
import { valueConfigs } from "./generated/valueConfigs";

export const devConfigs = {
  sheetConfigs,
  columnConfigs,
  valueConfigs,
};

declare module "../src/01_SpreadsheetSchema/configRegister" {
  interface Register {
    configs: typeof devConfigs;
  }
}
