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
        columnNamesLeftToRight,
      },
      endRowIndex: table.endRowIndex,
    };
  }
  // A handle a person changed since the last sync names no rule now; deleting by a guess could take the wrong one.
  noSuchRuleStops(removals: PlannedRuleRemoval[]): SyncNoSuchRuleStop[] {
    // Read before the send, apiTableColumnsByTableId can lack a Table whose column configs only the regen writes.
    if (removals.length === 0) return [];
    const apiTableColumnsByTableId =
      this.columnConfigOperator.apiTableColumnsByTableId();
    return removals.flatMap((removal): SyncNoSuchRuleStop[] => {
      // A Table a person deleted took its rules with it, so the planner drops the entry.
      const tableId = this.tableConfigOperator.tableIdOf(removal.table);
      if (!this.ss.raw.tableIdIsActive(tableId)) return [];
      const indexes = this._ruleIndexesByHandle(
        removal.table,
        apiTableColumnsByTableId,
      );
      if (indexes.has(removal.handle)) return [];
      return [
        {
          kind: "noSuchRule",
          table: removal.table,
          path: ruleRemovalLines.path(removal),
          handles: [...indexes.keys()],
        },
      ];
    });
  }
  // Matched on the same read the stops were, so every handle names a rule.
  removeRules(removals: PlannedRuleRemoval[]): this {
    if (removals.length === 0) return this;
    const apiTableColumnsByTableId =
      this.columnConfigOperator.apiTableColumnsByTableId();
    for (const removal of removals) {
      const index = Val.assert(
        this._ruleIndexesByHandle(removal.table, apiTableColumnsByTableId).get(
          removal.handle,
        ),
        `rule ${removal.handle} on ${removal.table}`,
      );
      this._sheetOf(removal.table).removeConditionalFormatRuleByIndex(index);
    }
    return this;
  }
  // The removals whose handle the refetch still finds on their Table.
  removalsDidntLand(removals: PlannedRuleRemoval[]): SyncSendLine[] {
    if (removals.length === 0) return [];
    const apiTableColumnsByTableId =
      this.columnConfigOperator.apiTableColumnsByTableId();
    return removals.flatMap((removal): SyncSendLine[] => {
      const indexes = this._ruleIndexesByHandle(
        removal.table,
        apiTableColumnsByTableId,
      );
      if (!indexes.has(removal.handle)) return [];
      const { subject, what } = ruleRemovalLines.sendLine(removal);
      return [{ subject, what: `${what}; Sheets kept it` }];
    });
  }
  // Named as the regen names them, so a handle matches the one the configs record.
  private _ruleIndexesByHandle(
    table: string,
    apiTableColumnsByTableId: Map<string, ApiTableColumns>,
  ): Map<string, number> {
    const tableId = this.tableConfigOperator.tableIdOf(table);
    return RuleRecords.ruleIndexesByHandle(
      this._ruleRecordsTab(this._sheetOf(table), apiTableColumnsByTableId),
      tableId,
    );
  }
  private _sheetOf(table: string): SheetRaw {
    return this.ss.raw.table(this.tableConfigOperator.tableIdOf(table)).sheet;
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
