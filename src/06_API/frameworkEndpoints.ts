import type { CodebaseNameDelimiter } from "../00_Source/CellValues/cellValues";
import { configTableNames } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import {
  isInTnGroup,
  type SheetNameWithIdColumn,
} from "../04_SpreadsheetNamed/TableNameGroups";
import type { Endpoint, Endpoints, EndpointsAll } from "./Endpoints";

// One per Table whose config has `hasIdColumn: true`, read at runtime, so nothing is generated for it.
export type FillRowIdsEndpoints = {
  [
    TN in SheetNameWithIdColumn as `${TN}${CodebaseNameDelimiter}id`
  ]: Endpoint<TN>;
};

export function withFrameworkEndpoints(endpoints: Endpoints): EndpointsAll {
  return { ...endpoints, ...fillRowIdsEndpoints() };
}

function fillRowIdsEndpoints(): FillRowIdsEndpoints {
  return configTableNames()
    .filter((tableName) => isInTnGroup("hasIdColumn", tableName))
    .reduce((acc, tableName) => {
      acc[`${tableName}_id`] = fillRowIdsEndpoint(tableName);
      return acc;
    }, {} as FillRowIdsEndpoints);
}

// No feedback columns, so the run reports by toast; flushed first, so the toast follows the IDs onto the sheet.
function fillRowIdsEndpoint(
  tableName: SheetNameWithIdColumn,
): Pick<Endpoint<SheetNameWithIdColumn>, "action"> {
  return {
    action: (ss) => {
      ss.fillMissingRowIds([tableName]);
      ss.batchUpdateGSheets();
      ss.toast(
        `Filled the blank IDs in "${ss.table(tableName).raw.sheet.title}".`,
      );
    },
  };
}
