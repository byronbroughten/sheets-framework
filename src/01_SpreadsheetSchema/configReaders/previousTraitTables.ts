import { type BraceConfigs, BraceFormula } from "./BraceFormula";
import type { ColumnConfigStored } from "./makeConfigs";
import type { TableReading, TraitReadings } from "./SyncPlanner";

// The previous side of the comparison, read from the configs the last run wrote.
export function previousTraitTables(
  configs: BraceConfigs,
): TraitReadings["tables"] {
  return Object.entries(configs.columnConfigs).reduce<TraitReadings["tables"]>(
    (tables, [tableKey, columns]) => {
      const tableName = configs.tableConfigs[tableKey]?.tableName;
      tables[tableKey] = {
        traits: tableName === undefined ? {} : { name: tableName },
        columns: Object.entries(columns).reduce<TableReading["columns"]>(
          (
            byName,
            [columnName, { header, emptyValueAllowed, formula, columnType }],
          ) => {
            byName[columnName] = {
              header,
              emptyValueAllowed,
              formula: formula ?? undefined,
              columnType: previousColumnType({ configs, tableKey, columnType }),
            };
            return byName;
          },
          {},
        ),
      };
      return tables;
    },
    {},
  );
}

interface PreviousColumnTypeProps {
  configs: BraceConfigs;
  tableKey: string;
  columnType: ColumnConfigStored["columnType"];
}

// A rule recorded as Sheets wrote it, before rules were braced, counts as not recorded, so live is the baseline.
function previousColumnType({
  configs,
  tableKey,
  columnType,
}: PreviousColumnTypeProps): ColumnConfigStored["columnType"] {
  const rule = columnType?.rule;
  if (rule === undefined || rule === null) return columnType;
  if (BraceFormula.isDropdownRuleInBraceForm({ configs, tableKey, rule })) {
    return columnType;
  }
  return undefined;
}
