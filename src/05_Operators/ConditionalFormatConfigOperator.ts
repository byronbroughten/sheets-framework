import { Val } from "@byronbroughten/utils/val";

import { columnConfigsByName } from "../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import {
  type ConditionalFormatConfigsBase,
  makeImportLine,
} from "../01_SpreadsheetSchema/configReaders/makeConfigs";
import {
  RuleRecords,
  type RuleRecordsTab,
  type RuleRecordsTable,
  type TableRuleIndex,
} from "../01_SpreadsheetSchema/configReaders/RuleRecords";
import {
  type PlannedRuleRemoval,
  ruleRemovalLines,
} from "../01_SpreadsheetSchema/configReaders/SyncPlanner";
import type {
  SyncNoSuchRuleStop,
  SyncPlanEffect,
  SyncSendLine,
} from "../01_SpreadsheetSchema/configReaders/syncRunResult";
import type { SheetRaw } from "../02_SpreadsheetRaw/SheetRaw";
import type { TableRaw } from "../02_SpreadsheetRaw/TableRaw";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import {
  type ApiTableColumns,
  ColumnConfigsOperator,
} from "./ColumnConfigsOperator";
import {
  columnNamesByColumnId,
  oneLinePerLeafFileSource,
  tablesAndTabsFileSource,
} from "./configFileSource";
import { SpreadsheetBaseOperator } from "./SpreadsheetBaseOperator";
import { TableConfigsOperator } from "./TableConfigsOperator";

// A Table a plan deletes: its codebase name and its ID.
export interface DeletedTable {
  table: string;
  tableId: string;
}

// A removal matched on the pre-send read; `copiesLeft` is how many rules with its shared handle the send leaves on the Table.
export interface IndexedRuleRemoval extends PlannedRuleRemoval, TableRuleIndex {
  tableId: string;
  copiesLeft: number;
}

interface TableRules {
  tableId: string;
  indexes: Map<string, TableRuleIndex>;
}

// Reads the rules the regen's plain property get brought into Raw, so it costs no round trip of its own.
export class ConditionalFormatConfigOperator extends SpreadsheetBaseOperator {
  static init(): ConditionalFormatConfigOperator {
    return new ConditionalFormatConfigOperator(
      SpreadsheetBaseOperator.initOperatorProps(),
    );
  }
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  get columnConfigOperator(): ColumnConfigsOperator {
    return new ColumnConfigsOperator(this.operatorProps);
  }
  get tableConfigOperator(): TableConfigsOperator {
    return new TableConfigsOperator(this.operatorProps);
  }
  newConditionalFormatConfigs(): ConditionalFormatConfigsBase {
    const apiTableColumnsByTableId =
      this.columnConfigOperator.apiTableColumnsByTableId();
    const { raw } = this.ss;
    const tabs = raw.activeSheetGids.map((sheetGid) =>
      this._ruleRecordsTab(raw.sheet(sheetGid), apiTableColumnsByTableId),
    );
    return RuleRecords.build(tabs);
  }
  private _ruleRecordsTab(
    sheet: SheetRaw,
    apiTableColumnsByTableId: Map<string, ApiTableColumns>,
  ): RuleRecordsTab {
    const tables = sheet.tableIds.flatMap((tableId) => {
      const apiTableColumns = apiTableColumnsByTableId.get(tableId);
      if (apiTableColumns === undefined) return [];
      return [this._ruleRecordsTable(tableId, apiTableColumns)];
    });
    return {
      tabName: this.schema.titleToName(sheet.title),
      sheetGid: sheet.sheetGid,
      tables,
      rules: sheet.conditionalFormatRules(),
    };
  }
  // Named from the previous configs: apiTableColumnsByTableId holds only after the regen, which no longer has the Table.
  deletionEffects({
    table: tableKey,
    tableId,
  }: DeletedTable): SyncPlanEffect[] {
    const columnConfigs = Val.assert(
      columnConfigsByName()[tableKey],
      `previous column configs of ${tableKey}`,
    );
    const { sheet } = this.ss.raw.table(tableId);
    return RuleRecords.deletionEffects({
      tabName: this.schema.titleToName(sheet.title),
      rules: sheet.conditionalFormatRules(),
      table: this._ruleRecordsTable(tableId, { tableKey, columnConfigs }),
    });
  }
  private _ruleRecordsTable(
    tableId: string,
    { tableKey, columnConfigs }: ApiTableColumns,
  ): RuleRecordsTable {
    const table = this.ss.raw.table(tableId);
    const columnNamesById = columnNamesByColumnId(columnConfigs);
    const columnNamesLeftToRight = table.fullTableColIndexes.map((colIndex) =>
      columnNameAt({ table, colIndex, columnNamesById }),
    );
    return {
      tableKey,
      tableId,
      layout: {
        headerTopLeft: {
          rowIndex: table.startRowIndex,
          colIndex: table.startColumnIndex,
        },
        endRowIndex: table.endRowIndex,
        columnNamesLeftToRight,
      },
    };
  }
  // A handle a person changed since the last sync names no rule now; deleting by a guess could take the wrong one.
  noSuchRuleStops(removals: PlannedRuleRemoval[]): SyncNoSuchRuleStop[] {
    const rulesByTable = this._tableRulesByTable(removals);
    return removals.flatMap((removal): SyncNoSuchRuleStop[] => {
      // A Table a person deleted took its rules with it, so the planner drops the entry.
      const rules = rulesByTable.get(removal.table);
      if (rules === undefined || rules.indexes.has(removal.handle)) return [];
      return [
        {
          kind: "noSuchRule",
          table: removal.table,
          path: ruleRemovalLines.path(removal),
          handles: [...rules.indexes.keys()],
        },
      ];
    });
  }
  // Indexed on the read the stops were matched on, so every handle names a rule.
  indexedRuleRemovals(removals: PlannedRuleRemoval[]): IndexedRuleRemoval[] {
    const rulesByTable = this._tableRulesByTable(removals);
    return removals.map((removal) => {
      const { tableId, indexes } = Val.assert(
        rulesByTable.get(removal.table),
        `rules of ${removal.table}`,
      );
      const { sharedHandle, index } = Val.assert(
        indexes.get(removal.handle),
        `rule ${removal.handle} on ${removal.table}`,
      );
      const removedCopies = removals.filter(
        (other) =>
          other.table === removal.table &&
          indexes.get(other.handle)?.sharedHandle === sharedHandle,
      ).length;
      return {
        ...removal,
        tableId,
        index,
        sharedHandle,
        copiesLeft: copyCount(indexes, sharedHandle) - removedCopies,
      };
    });
  }
  removeRules(removals: IndexedRuleRemoval[]): this {
    for (const { tableId, index } of removals) {
      this.ss.raw
        .table(tableId)
        .sheet.removeConditionalFormatRuleByIndex(index);
    }
    return this;
  }
  // By copies, not handle: removing one of two equal rules renames the other to the removed one's handle.
  removalsDidntLand(removals: IndexedRuleRemoval[]): SyncSendLine[] {
    const rulesByTable = this._tableRulesByTable(removals);
    return removals.flatMap((removal): SyncSendLine[] => {
      const indexes = rulesByTable.get(removal.table)?.indexes ?? new Map();
      if (copyCount(indexes, removal.sharedHandle) <= removal.copiesLeft) {
        return [];
      }
      const { subject, what } = ruleRemovalLines.sendLine(removal);
      return [{ subject, what: `${what}; Sheets kept it` }];
    });
  }
  // Named as the regen names them, so a handle matches the one the configs record; one tab walk per Table.
  private _tableRulesByTable(
    removals: PlannedRuleRemoval[],
  ): Map<string, TableRules> {
    const rulesByTable = new Map<string, TableRules>();
    if (removals.length === 0) return rulesByTable;
    const apiTableColumnsByTableId =
      this.columnConfigOperator.identifiedApiTableColumnsByTableId();
    for (const table of new Set(removals.map((removal) => removal.table))) {
      const tableId = this.tableConfigOperator.tableIdOf(table);
      if (!this.ss.raw.tableIdIsActive(tableId)) continue;
      const tab = this._ruleRecordsTab(
        this.ss.raw.table(tableId).sheet,
        apiTableColumnsByTableId,
      );
      rulesByTable.set(table, {
        tableId,
        indexes: RuleRecords.ruleIndexesByHandle(tab, tableId),
      });
    }
    return rulesByTable;
  }
  toFileSource(): string {
    return [
      makeImportLine(["makeConditionalFormatConfigs"]),
      ``,
      `export const conditionalFormatConfigs = makeConditionalFormatConfigs(${tablesAndTabsFileSource(
        this.newConditionalFormatConfigs(),
        oneLinePerLeafFileSource,
      )});`,
      ``,
    ].join("\n");
  }
}

function copyCount(
  indexes: Map<string, TableRuleIndex>,
  sharedHandle: string,
): number {
  return [...indexes.values()].filter(
    (rule) => rule.sharedHandle === sharedHandle,
  ).length;
}

interface ColumnNameAtProps {
  table: TableRaw;
  colIndex: number;
  columnNamesById: Map<string, string>;
}

// Undefined for a column with no ID: a run that mints none leaves it out of the configs.
function columnNameAt({
  table,
  colIndex,
  columnNamesById,
}: ColumnNameAtProps): string | undefined {
  const columnId = table.columnResolver.columnIdAt(colIndex);
  if (columnId === "") return undefined;
  return Val.assert(
    columnNamesById.get(columnId),
    `column name at column index ${colIndex} of ${table.name}`,
  );
}
