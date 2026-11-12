import {
  type CodebaseNameDelimiter,
  codebaseNameDelimiter,
} from "../00_Source/CellValues/cellValues";
import {
  type SheetNameWithIdColumn,
  tableNamesWithIdColumn,
} from "../04_SpreadsheetNamed/TableNameGroups";
import type { Endpoint, Endpoints, EndpointsAll } from "./Endpoints";

// Not MakeColumnFullName: "id" doesn't satisfy its ColumnName<TN> constraint over a generic TN.
export type FillRowIdsEndpoints = {
  [
    TN in SheetNameWithIdColumn as `${TN}${CodebaseNameDelimiter}id`
  ]: Endpoint<TN>;
};

export function withFrameworkEndpoints(endpoints: Endpoints): EndpointsAll {
  return { ...endpoints, ...fillRowIdsEndpoints() };
}

function fillRowIdsEndpoints(): FillRowIdsEndpoints {
  return tableNamesWithIdColumn().reduce((acc, tableName) => {
    acc[`${tableName}${codebaseNameDelimiter}id`] =
      fillRowIdsEndpoint(tableName);
    return acc;
  }, {} as FillRowIdsEndpoints);
}

// Flushed before the toast, so the toast follows the IDs onto the sheet.
function fillRowIdsEndpoint(
  tableName: SheetNameWithIdColumn,
): Pick<Endpoint<SheetNameWithIdColumn>, "action"> {
  return {
    action: (ss) => {
      ss.fillMissingRowIds([tableName]);
      ss.batchUpdateGSheets();
      ss.toast(
        `Filled the blank IDs in "${ss.table(tableName).raw.sheetTitle}".`,
      );
    },
  };
}
