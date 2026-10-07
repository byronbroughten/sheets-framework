import { getColumnTraitByName } from "../01_SpreadsheetSchema/columnConfigsTypes";
import {
  configSheetFloorSeed,
  floorSeedColumnById,
} from "../01_SpreadsheetSchema/configSheetFloorSeed";
import { dimensionIds } from "../01_SpreadsheetSchema/dimensionIds";
import {
  type ColumnConfigsGeneric,
  makeImportLine,
} from "../01_SpreadsheetSchema/makeConfigs";
import { type ValueName } from "../01_SpreadsheetSchema/valueSchemas";
import type { ColumnProfileRaw } from "../02_SpreadsheetRaw/ColumnProfileRaw";
import { Str } from "../utils/Str";
import { columnConfigsFileSource } from "./configFileSource";
import { GenericTableOperator } from "./GenericTableOperator";
import {
  type ConfigSyncState,
  type OperatorProps,
  SpreadsheetBaseOperator,
  type UntypedHeadersByTableName,
} from "./SpreadsheetBaseOperator";
import { TableConfigOperator } from "./TableConfigOperator";
import { ValueConfigOperator } from "./ValueConfigOperator";

interface ColumnIdentity {
  tableId: string;
  columnId: string;
}

export class ColumnConfigOperator extends GenericTableOperator<"columnConfig"> {
  constructor(props: OperatorProps) {
    super({
      tableName: "columnConfig",
      ...props,
    });
  }
  static init(): ColumnConfigOperator {
    return new ColumnConfigOperator(
      SpreadsheetBaseOperator.initOperatorProps(),
    );
  }
  get columnConfigSync(): ConfigSyncState["columnConfigSync"] {
    return this.configSyncState.columnConfigSync;
  }
  get tableConfigOperator(): TableConfigOperator {
    return new TableConfigOperator(this.operatorProps);
  }
  get valueConfigOperator(): ValueConfigOperator {
    return new ValueConfigOperator(this.operatorProps);
  }
  private get untypedHeadersByTableName(): UntypedHeadersByTableName {
    return this.columnConfigSync.untypedHeadersByTableName;
  }
  // Derived fresh each call, not cached — a stored field goes stale across
  // this coordinator's per-access getter rebuilds.
  private get tableIdsApiAccesses(): Set<string> {
    return new Set(this.tableConfigOperator.tableIdsApiAccesses());
  }
  activeValueTitles(): string[] {
    return this.table.workingRowIndexesWithData.map((rowIndex) =>
      this._describedColumn(this._columnIdentity(rowIndex)).valueTitle(),
    );
  }
  assertSyncedToSpreadsheet(): void {
    if (!this.columnConfigSync.syncedToSpreadsheet) {
      throw new Error(
        "ColumnConfigOperator has not yet synced to the spreadsheet.",
      );
    }
  }
  prepFetchWithTableConfig(): void {
    this.tableConfigOperator.assertPrepFetchIsComplete();
    this.table.prepFetchColumnsFull(
      "tableId",
      "columnId",
      "tableName",
      "header",
      "emptyValueAllowed",
    );
  }
  fetchAfterTableConfigSynced(): this {
    this.tableConfigOperator.assertSyncedToSpreadsheet();
    this.tableIdsApiAccesses.forEach((tableId) => {
      const table = this.ss.raw.table(tableId);
      // hasIdColumn samples this row after Let api access is known.
      table.headRow("header").gatherFetchFull();
      table.headRow("columnId").gatherFetchFull();
      table.topRow.gatherFetchFull();
    });
    this.ss.raw.fetchAllGathered(true);
    return this;
  }
  syncToSpreadsheet(): this {
    this._addMissingColumnIds();
    this._pruneColumnRows();
    this._appendColumnRows();
    this._updateProgrammaticValues();
    this._logUntypedColumns();
    this.columnConfigSync.syncedToSpreadsheet = true;
    return this;
  }
  // Undefined, not "", so a fully typed spreadsheet still reports "Succeeded".
  untypedColumnsSummary(): string | undefined {
    this.assertSyncedToSpreadsheet();
    const untypedHeaders = Array.from(this.untypedHeadersByTableName.values());
    if (untypedHeaders.length === 0) {
      return undefined;
    }
    const columnCount = untypedHeaders.reduce(
      (count, headers) => count + headers.length,
      0,
    );
    const sentences = [
      `Succeeded, but ${columnCount} column(s) across ${untypedHeaders.length} ` +
        `Table(s) are untyped, so their value names were guessed. See the ` +
        `execution log for the list.`,
    ];
    const blankSampleNames = this._blankSampleTableNames();
    if (blankSampleNames.length > 0) {
      const names = blankSampleNames.map((name) => `"${name}"`).join(", ");
      sentences.push(
        `On ${blankSampleNames.length} of those Table(s) the top data row ` +
          `was blank, so the guess had no sample behind it: ${names}.`,
      );
    }
    return sentences.join(" ");
  }
  declaredCellReport(): string | undefined {
    const lines = this.columnConfigSync.declaredCellReportLines;
    if (lines.length === 0) return undefined;
    return lines.join(" ");
  }
  // Derived, since blank facts are deliberately indistinguishable from real ones.
  private _blankSampleTableNames(): string[] {
    const names: string[] = [];
    this.tableIdsApiAccesses.forEach((tableId) => {
      const table = this.ss.raw.table(tableId);
      if (!this.untypedHeadersByTableName.has(table.name)) return;
      if (!table.topDataRowIsBlank()) return;
      names.push(table.name);
    });
    return names;
  }
  private _addMissingColumnIds(): this {
    let idsAdded = 0;
    this.tableIdsApiAccesses.forEach((tableId) => {
      const idPrefix = this.tableConfigOperator.idPrefix(tableId);
      // Raw, not TableIdentified: a Table here may predate its generated schema.
      const table = this.ss.raw.table(tableId);
      const colIndexes = table.columnResolver.colIndexesWithoutColumnId;
      colIndexes.forEach((colIndex) => {
        table
          .headRow("columnId")
          .updateValue(colIndex, dimensionIds.col(idPrefix));
      });
      idsAdded += colIndexes.length;
    });
    Logger.log(
      `ensureColumnIds: prepared to add ${idsAdded} missing column ID(s)`,
    );
    return this;
  }
  private _pruneColumnRows(): this {
    const col = this.table.columns("tableId", "columnId");
    let staleCount = 0;
    this.table.workingRowIndexes.forEach((rowIndex) => {
      const tableId = col.tableId.valueOrEmpty(rowIndex);
      const columnId = col.columnId.valueOrEmpty(rowIndex);
      if (
        tableId === "" ||
        columnId === "" ||
        !this.tableIdsApiAccesses.has(tableId) ||
        !this._hasColumnId(tableId, columnId)
      ) {
        this.table.row(rowIndex).delete();
        staleCount++;
      }
    });
    Logger.log(`_pruneColumnRows: pruned ${staleCount} stale row(s).`);
    return this;
  }
  private _hasColumnId(tableId: string, columnId: string): boolean {
    return this.ss.raw.table(tableId).columnResolver.hasColumnId(columnId);
  }
  private _appendColumnRows(): this {
    const existingIdentityKeys = new Set(
      this.table.workingRowIndexesWithData.map((rowIndex) =>
        columnIdentityKey(this._columnIdentity(rowIndex)),
      ),
    );

    let appendedCount = 0;
    this.tableIdsApiAccesses.forEach((tableId) => {
      const { columnIds } = this.ss.raw.table(tableId).profile;
      columnIds.forEach((columnId) => {
        if (
          !existingIdentityKeys.has(columnIdentityKey({ tableId, columnId }))
        ) {
          this.table.appendRowWithVals({ tableId, columnId });
          appendedCount++;
        }
      });
    });
    Logger.log(`_appendColumnRows: added ${appendedCount} new row(s).`);
    return this;
  }
  private _columnIdentity(rowIndex: number): ColumnIdentity {
    const col = this.table.columns("tableId", "columnId");
    return {
      tableId: col.tableId.value(rowIndex),
      columnId: col.columnId.value(rowIndex),
    };
  }
  private _describedColumn({
    tableId,
    columnId,
  }: ColumnIdentity): ColumnProfileRaw {
    return this.ss.raw.table(tableId).profile.columnById(columnId);
  }
  private _updateProgrammaticValues(): void {
    const col = this.table.columns("tableName", "header", "emptyValueAllowed");
    const reportLines = this.columnConfigSync.declaredCellReportLines;
    reportLines.length = 0;
    let updatedValues = 0;
    this.untypedHeadersByTableName.clear();
    this.table.workingRowIndexesWithData.forEach((rowIndex) => {
      const identity = this._columnIdentity(rowIndex);

      const actualTableName = this.ss.raw.table(identity.tableId).name;
      if (col.tableName.valueOrEmpty(rowIndex) !== actualTableName) {
        col.tableName.cell(rowIndex).updateValue(actualTableName);
        updatedValues++;
      }

      const describedColumn = this._describedColumn(identity);
      const actualHeader = describedColumn.header;
      if (col.header.valueOrEmpty(rowIndex) !== actualHeader) {
        col.header.cell(rowIndex).updateValue(actualHeader);
        updatedValues++;
      }

      if (
        this._updateSelfDescribingEmptyValueAllowed({
          rowIndex,
          identity,
          tableName: actualTableName,
          header: actualHeader,
        })
      ) {
        updatedValues++;
      }

      if (describedColumn.declaredValueTitle() === undefined) {
        this._recordUntypedColumn(actualTableName, actualHeader);
      }
    });
    Logger.log(`Corrected ${updatedValues} inaccurate Column Config cell(s).`);
  }
  private _updateSelfDescribingEmptyValueAllowed({
    rowIndex,
    identity,
    tableName,
    header,
  }: {
    rowIndex: number;
    identity: ColumnIdentity;
    tableName: string;
    header: string;
  }): boolean {
    const seedColumn = floorSeedColumnById(identity.tableId, identity.columnId);
    if (seedColumn === undefined) return false;
    const emptyValueAllowed = this.table.column("emptyValueAllowed");
    if (
      emptyValueAllowed.valueOrEmpty(rowIndex) === seedColumn.emptyValueAllowed
    ) {
      return false;
    }
    emptyValueAllowed.cell(rowIndex).updateValue(seedColumn.emptyValueAllowed);
    this.columnConfigSync.declaredCellReportLines.push(
      `${configSheetFloorSeed.columnConfig.title} · ${getColumnTraitByName(
        "columnConfig",
        "emptyValueAllowed",
        "header",
      )} · ${tableName} · ${header} → ${
        seedColumn.emptyValueAllowed ? "TRUE" : "FALSE"
      }`,
    );
    return true;
  }
  private _recordUntypedColumn(tableName: string, header: string): void {
    const headers = this.untypedHeadersByTableName.get(tableName) ?? [];
    headers.push(header);
    this.untypedHeadersByTableName.set(tableName, headers);
  }
  private _logUntypedColumns(): void {
    this.untypedHeadersByTableName.forEach((headers, tableName) => {
      Logger.log(
        `Untyped columns on Table "${tableName}" (${headers.length}): ${headers.join(", ")}`,
      );
    });
  }
  newColumnConfigs(): ColumnConfigsGeneric {
    const tableKeysByTableId = this.tableConfigOperator.tableKeysByTableId();
    const col = this.table.columns("header", "emptyValueAllowed");
    const columnConfigs: ColumnConfigsGeneric = {};
    this.table.workingRowIndexesWithData.forEach((rowIndex) => {
      const identity = this._columnIdentity(rowIndex);
      const { tableId, columnId } = identity;
      const header = col.header.value(rowIndex);
      const tableKey = tableKeysByTableId.get(tableId);
      if (tableKey === undefined) {
        throw new Error(
          `generateColumnConfigFileSource: column "${columnId}" references Table ID ` +
            `"${tableId}", which has no ticked row in Table Config.`,
        );
      }
      const columnName = Str.sentenceToCamelCase(header);
      if (!columnConfigs[tableKey]) {
        columnConfigs[tableKey] = {};
      }
      const tableColumnConfigs = columnConfigs[tableKey];
      if (tableColumnConfigs[columnName]) {
        throw new Error(
          `generateColumnConfigFileSource: duplicate column name "${columnName}" ` +
            `derived from header "${header}" on Table "${tableKey}".`,
        );
      }
      const describedColumn = this._describedColumn(identity);
      tableColumnConfigs[columnName] = {
        columnId,
        header,
        valueName: this.schema.titleToName(
          describedColumn.valueTitle(),
        ) as ValueName,
        isFormula: describedColumn.isFormula,
        emptyValueAllowed: col.emptyValueAllowed.value(rowIndex),
        customDefaultValue: null,
      };
    });
    return columnConfigs;
  }
  toFileSource(makeConfigsImport: string): string {
    return [
      `${makeImportLine("makeColumnConfigs", makeConfigsImport)}`,
      ``,
      `export const columnConfigs = makeColumnConfigs(${columnConfigsFileSource(
        this.newColumnConfigs(),
      )});`,
      ``,
    ].join("\n");
  }
}

function columnIdentityKey({ tableId, columnId }: ColumnIdentity): string {
  return `${tableId}:${columnId}`;
}
