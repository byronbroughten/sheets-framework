import type { configSheetFloorSeed } from "../01_SpreadsheetSchema/configReaders/configSheetFloorSeed";
import type { Endpoint, Endpoints, EndpointsAll } from "./Endpoints";

// The floor seed still lists Sync Config until the floor goes.
type SeededFrameworkEndpoints = {
  [
    K in Exclude<
      keyof typeof configSheetFloorSeed.spreadsheetConfig.endpoints,
      "spreadsheetConfig_syncConfigSheetRowsTimeLastRan"
    >
  ]: Endpoint<"spreadsheetConfig">;
};

export const frameworkEndpoints = {
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
