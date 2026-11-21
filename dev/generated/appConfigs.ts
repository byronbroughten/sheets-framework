// Without it, a program that reaches this file only by dynamic import rejects the augmentation (TS2664).
import type {} from "@byronbroughten/sheets-framework";

import { columnConfigs } from "./columnConfigs.ts";
import { conditionalFormatConfigs } from "./conditionalFormatConfigs.ts";
import type { IdNameMap } from "./idNameMap.ts";
import { retiredIdPrefixes } from "./retiredIdPrefixes.ts";
import { tabConfigs } from "./tabConfigs.ts";
import { tableConfigs } from "./tableConfigs.ts";
import { valueConfigs } from "./valueConfigs.ts";

export const appConfigs = {
  tableConfigs,
  tabConfigs,
  conditionalFormatConfigs,
  retiredIdPrefixes,
  columnConfigs,
  valueConfigs,
};

declare module "@byronbroughten/sheets-framework" {
  interface Register {
    configs: typeof appConfigs;
    idNameMap: IdNameMap;
  }
}
