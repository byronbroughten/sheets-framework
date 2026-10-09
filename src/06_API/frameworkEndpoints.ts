import type { configSheetFloorSeed } from "../01_SpreadsheetSchema/configReaders/configSheetFloorSeed";
import { ConfigCoordinator } from "../05_Operators/ConfigCoordinator";
import type { Endpoint, Endpoints, EndpointsAll } from "./Endpoints";

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

export function withFrameworkEndpoints(endpoints: Endpoints): EndpointsAll {
  return { ...endpoints, ...frameworkEndpoints };
}
