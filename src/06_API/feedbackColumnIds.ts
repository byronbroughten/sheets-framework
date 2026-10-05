import type { TableName } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import { SpreadsheetSchema } from "../01_SpreadsheetSchema/SpreadsheetSchema";
import {
  type FeedbackColumnIds,
  installFeedbackColumnIds,
} from "../03_SpreadsheetIdentified/feedbackColumnRegister";
import type { Endpoints, EndpointsAll } from "./Endpoints";
import { withFrameworkEndpoints } from "./frameworkEndpoints";

// Throws before installConfigs, since resolving a feedback column reads the configs.
export function installEndpoints(endpoints: Endpoints): void {
  installFeedbackColumnIds(
    feedbackColumnIdsOf(withFrameworkEndpoints(endpoints)),
  );
}

// Every endpoint's, not just the running one's, so an append reuses a row any endpoint stamped.
export function feedbackColumnIdsOf(
  endpoints: EndpointsAll,
): FeedbackColumnIds {
  const schema = new SpreadsheetSchema();
  return Object.entries(endpoints).reduce((acc, [fullName, endpoint]) => {
    const sheet = schema.sheetByColumnFullName(fullName);
    const columnIds = acc.get(sheet.sheetName) ?? new Set<string>();
    for (const columnName of [endpoint.timeLastRan, endpoint.runStatus]) {
      if (columnName === undefined) continue;
      columnIds.add(sheet.columnByName(columnName).columnId);
    }
    // No empty entries, so two maps declaring the same feedback columns install as the same set.
    if (columnIds.size > 0) acc.set(sheet.sheetName, columnIds);
    return acc;
  }, new Map<TableName, Set<string>>());
}
