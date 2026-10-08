// Without it, a program that reaches this file only by dynamic import rejects the augmentation (TS2664).
import type {} from "@byronbroughten/sheets-framework";

import { columnConfigs } from "./columnConfigs";
import { tableConfigs } from "./tableConfigs";
import { valueConfigs } from "./valueConfigs";

export const appConfigs = {
  tableConfigs,
  columnConfigs,
  valueConfigs,
};

declare module "@byronbroughten/sheets-framework" {
  interface Register {
    configs: typeof appConfigs;
  }
}
