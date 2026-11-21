import {
  type BraceConfigs,
  BraceFormula,
  type ColumnTypeProps,
} from "./BraceFormula.ts";
import type { ColumnConfigStored } from "./makeConfigs.ts";
import type { TableReading, TraitReadings } from "./SyncPlanner.ts";

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

// A rule recorded as Sheets wrote it, before rules were braced, counts as not recorded, so live is the baseline.
function previousColumnType(
  props: ColumnTypeProps,
): ColumnConfigStored["columnType"] {
  const recorded = props.columnType?.rule?.values ?? [];
  const braced = BraceFormula.decompileColumnType(props)?.rule?.values ?? [];
  const isInBraceForm = braced.every(
    (value, index) => value === recorded[index],
  );
  if (!isInBraceForm) return undefined;
  return props.columnType;
}
