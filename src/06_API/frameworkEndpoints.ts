import type { configSheetFloorSeed } from "../01_SpreadsheetSchema/configSheetFloorSeed";
import { ConfigCoordinator } from "../05_Operators/ConfigCoordinator";
import type { Endpoint } from "./Endpoints";

type SeededFrameworkEndpoints = {
  [
    K in keyof typeof configSheetFloorSeed.spreadsheetConfig.endpoints
  ]: Endpoint<"spreadsheetConfig">;
};

export const frameworkEndpoints = {
  spreadsheetConfig_syncConfigSheetRowsTimeLastRan: {
    action: (ss) =>
      new ConfigCoordinator(ss.spreadsheetNamedProps).syncConfigSheetRows(),
    timeLastRan: "syncConfigSheetRowsTimeLastRan",
    runStatus: "syncConfigSheetRowsRunStatus",
  },
  spreadsheetConfig_fillRowIdsTimeLastRan: {
    action: (ss) => {
      ss.fillMissingRowIds();
    },
    timeLastRan: "fillRowIdsTimeLastRan",
    runStatus: "fillRowIdsRunStatus",
  },
} as const satisfies SeededFrameworkEndpoints;
