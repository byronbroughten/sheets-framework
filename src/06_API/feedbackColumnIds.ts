import { SpreadsheetSchema } from "../01_SpreadsheetSchema/SpreadsheetSchema";
import type { FeedbackColumnIds } from "../03_SpreadsheetIdentified/ClassBases/SpreadsheetBaseIdentified";
import type { EndpointsAll } from "./Endpoints";

// Every endpoint's, not just the running one's, so an append reuses a row any endpoint stamped.
export function feedbackColumnIdsOf(
  endpoints: EndpointsAll,
): FeedbackColumnIds {
  const schema = new SpreadsheetSchema();
  return Object.entries(endpoints).reduce((acc, [fullName, endpoint]) => {
    const sheet = schema.sheetByColumnFullName(fullName);
    const columnIds = acc.get(sheet.sheetGid) ?? new Set<string>();
    for (const columnName of [endpoint.timeLastRan, endpoint.runStatus]) {
      if (columnName === undefined) continue;
      columnIds.add(sheet.columnByName(columnName).columnId);
    }
    acc.set(sheet.sheetGid, columnIds);
    return acc;
  }, new Map<number, Set<string>>());
}
