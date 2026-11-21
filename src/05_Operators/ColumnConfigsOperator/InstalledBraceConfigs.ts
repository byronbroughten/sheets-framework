import {
  type BraceConfigs,
  BraceFormula,
} from "../../01_SpreadsheetSchema/configReaders/BraceFormula";
import { columnConfigsByName } from "../../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import type { ColumnConfigStored } from "../../01_SpreadsheetSchema/configReaders/makeConfigs";
import type { IntendedTrait } from "../../01_SpreadsheetSchema/configReaders/SyncPlanner";
import { tableConfigsByName } from "../../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import type { SpreadsheetRaw } from "../../02_SpreadsheetRaw/SpreadsheetRaw";
import type { TableRaw } from "../../02_SpreadsheetRaw/TableRaw";

// The installed configs, and the names a formula or rule is braced or compiled against: each the installed configs with Table names swapped.
export class InstalledBraceConfigs {
  readonly installed: BraceConfigs = {
    tableConfigs: tableConfigsByName(),
    columnConfigs: columnConfigsByName(),
  };
  // The previous configs with each API-access Table's live name and each live column's header: live names under previous keys, the keys an alteration uses, so a person renaming a Table a formula or rule names reads as no change.
  // Takes its caller's raw and Table IDs: it shares _withTableNames with atColumnFill here (#345).
  liveNamed(
    raw: SpreadsheetRaw,
    apiTableIds: ReadonlySet<string>,
  ): BraceConfigs {
    const apiTables = Object.entries(this.installed.tableConfigs).reduce(
      (byKey, [tableKey, { tableId }]) => {
        if (apiTableIds.has(tableId)) byKey.set(tableKey, raw.table(tableId));
        return byKey;
      },
      new Map<string, TableRaw>(),
    );
    const named = this._withTableNames(
      new Map(
        [...apiTables].map(([tableKey, table]) => [tableKey, table.name]),
      ),
    );
    return Object.keys(named.tableConfigs).reduce<BraceConfigs>(
      (live, tableKey) => {
        const previousColumns = named.columnConfigs[tableKey] ?? {};
        const table = apiTables.get(tableKey);
        if (table === undefined) {
          live.columnConfigs[tableKey] = previousColumns;
          return live;
        }
        live.columnConfigs[tableKey] = liveHeaders(table, previousColumns);
        return live;
      },
      { tableConfigs: named.tableConfigs, columnConfigs: {} },
    );
  }
  // The names on the sheet when a column fill or column-property set lands: Table renames go out before both and header writes after (docs/architecture/flush-order.md).
  atColumnFill(sent: IntendedTrait[]): BraceConfigs {
    return this._withTableNames(
      new Map(
        sent.flatMap((intended): [string, string][] => {
          if (intended.kind !== "table" || intended.trait !== "name") return [];
          return [[intended.table, String(intended.value)]];
        }),
      ),
    );
  }
  private _withTableNames(
    tableNamesByKey: ReadonlyMap<string, string>,
  ): BraceConfigs {
    const { tableConfigs, columnConfigs } = this.installed;
    return {
      tableConfigs: Object.entries(tableConfigs).reduce<
        BraceConfigs["tableConfigs"]
      >((byKey, [tableKey, tableConfig]) => {
        const tableName = tableNamesByKey.get(tableKey);
        if (tableName === undefined) {
          byKey[tableKey] = tableConfig;
          return byKey;
        }
        byKey[tableKey] = { ...tableConfig, tableName };
        return byKey;
      }, {}),
      columnConfigs,
    };
  }
}

function liveHeaders(
  table: TableRaw,
  previousColumns: BraceConfigs["columnConfigs"][string],
): BraceConfigs["columnConfigs"][string] {
  const liveColumnIds = new Set(table.profile.columnIds);
  return Object.entries(previousColumns).reduce<
    BraceConfigs["columnConfigs"][string]
  >((byName, [columnName, columnConfig]) => {
    if (!liveColumnIds.has(columnConfig.columnId)) return byName;
    byName[columnName] = {
      ...columnConfig,
      header: table.profile.columnById(columnConfig.columnId).header,
    };
    return byName;
  }, {});
}

export function bracedFormula({
  configs,
  tableKey,
  formula,
}: {
  configs: BraceConfigs;
  tableKey: string;
  formula: ColumnConfigStored["formula"];
}): ColumnConfigStored["formula"] {
  if (formula === undefined || formula === null) return formula;
  return BraceFormula.decompile({ configs, tableKey, formula });
}
