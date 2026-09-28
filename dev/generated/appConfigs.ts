// Without it, a program that reaches this file only by dynamic import rejects the augmentation (TS2664).
import type {} from "@byronbroughten/sheets-framework";

import { columnConfigs } from "./columnConfigs";
import { sheetConfigs } from "./sheetConfigs";
import { valueConfigs } from "./valueConfigs";

export const appConfigs = { sheetConfigs, columnConfigs, valueConfigs };

declare module "@byronbroughten/sheets-framework" {
  interface Register {
    configs: typeof appConfigs;
  }
}
