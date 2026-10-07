import { getColumnTraitByName } from "../01_SpreadsheetSchema/columnConfigsTypes";
import {
  configSheetFloorSeed,
  floorTabSeedByTableId,
} from "../01_SpreadsheetSchema/configSheetFloorSeed";
import {
  idPrefixes,
  type IdPrefixLabel,
} from "../01_SpreadsheetSchema/idPrefixes";
import {
  makeImportLine,
  type TableConfigsBase,
} from "../01_SpreadsheetSchema/makeConfigs";
import { tableConfigsByTableId } from "../01_SpreadsheetSchema/tableConfigsTypes";
import type { TableRaw } from "../02_SpreadsheetRaw/TableRaw";
import { Val } from "../utils/Val";
import { oneLinePerEntryFileSource } from "./configFileSource";
import { GenericTableOperator } from "./GenericTableOperator";
import {
  type ConfigSyncState,
  type OperatorProps,
  SpreadsheetBaseOperator,
} from "./SpreadsheetBaseOperator";

export interface ColumnReference {
  tableId: string;
  columnId: string;
}

export class TableConfigOperator extends GenericTableOperator<"tableConfig"> {
  constructor(props: OperatorProps) {
    super({
      tableName: "tableConfig",
      ...props,
    });
  }
  static init(): TableConfigOperator {
    return new TableConfigOperator(SpreadsheetBaseOperator.initOperatorProps());
  }
  get tableConfigSync(): ConfigSyncState["tableConfigSync"] {
    return this.configSyncState.tableConfigSync;
  }
  assertPrepFetchIsComplete(): void {
    if (!this.tableConfigSync.prepFetchIsComplete) {
      throw new Error(
        "TableConfigOperator has not yet completed its prepFetch operation.",
      );
    }
  }
  assertSyncedToSpreadsheet(): void {
    if (!this.tableConfigSync.syncedToSpreadsheet) {
      throw new Error(
        "TableConfigOperator has not yet synced to the spreadsheet.",
      );
    }
  }
  prepFetchForSync(): void {
    this.table.prepFetchColumnsFull(
      "tableId",
      "tableName",
      "sheetTitle",
      "letApiAccess",
    );
    this.tableConfigSync.prepFetchIsComplete = true;
  }
  syncToSpreadsheet(): void {
    this._deleteStaleTableConfigs();
    this._appendMissingTableConfigs();
    this._updateProgrammaticValues();
    this.tableConfigSync.syncedToSpreadsheet = true;
  }
  private _deleteStaleTableConfigs(): void {
    this.table.rows.forEach((row) => {
      const tableId = row.valueOrEmpty("tableId");
      if (tableId === "" || !this.ss.raw.tableIdIsActive(tableId)) {
        row.delete();
      }
    });
  }
  private _appendMissingTableConfigs(): void {
    const colTableId = this.table.column("tableId");
    this.ss.raw.activeTableIds.forEach((tableId) => {
      if (!colTableId.hasValue(tableId)) {
        this.table.appendRowWithVals({ tableId });
      }
    });
  }
  private _updateProgrammaticValues(): void {
    const col = this.table.columns("tableId", "tableName", "sheetTitle");
    const reportLines = this.tableConfigSync.declaredCellReportLines;
    reportLines.length = 0;
    let updatedValues = 0;
    this.table.rowIndexesActiveWithData.forEach((rowIndex) => {
      const liveTable = this.ss.raw.table(col.tableId.value(rowIndex));
      if (col.tableName.valueOrEmpty(rowIndex) !== liveTable.name) {
        col.tableName.cell(rowIndex).updateValue(liveTable.name);
        updatedValues++;
      }
      if (col.sheetTitle.valueOrEmpty(rowIndex) !== liveTable.sheetTitle) {
        col.sheetTitle.cell(rowIndex).updateValue(liveTable.sheetTitle);
        updatedValues++;
      }
      if (this._updateSelfDescribingLetApiAccess(rowIndex, liveTable)) {
        updatedValues++;
      }
    });
    Logger.log(`Corrected ${updatedValues} inaccurate Table Config cells.`);
  }
  private _updateSelfDescribingLetApiAccess(
    rowIndex: number,
    liveTable: TableRaw,
  ): boolean {
    const seed = floorTabSeedByTableId(liveTable.tableId);
    if (seed === undefined) return false;
    const letApiAccess = this.table.column("letApiAccess");
    if (letApiAccess.valueOrEmpty(rowIndex) === seed.letApiAccess) return false;
    letApiAccess.cell(rowIndex).updateValue(seed.letApiAccess);
    this.tableConfigSync.declaredCellReportLines.push(
      `${configSheetFloorSeed.tableConfig.title} · ${getColumnTraitByName(
        "tableConfig",
        "letApiAccess",
        "header",
      )} · ${liveTable.sheetTitle} → ${seed.letApiAccess ? "TRUE" : "FALSE"}`,
    );
    return true;
  }
  tableIdsApiAccesses(): string[] {
    const col = this.table.columns("tableId", "letApiAccess");
    return this.table.rowIndexesActiveWithData.flatMap((rowIndex) => {
      if (!col.letApiAccess.valueOrEmpty(rowIndex)) return [];
      return [col.tableId.value(rowIndex)];
    });
  }
  idPrefix(tableId: string): string {
    return Val.assert(this._idPrefixesByTableId().get(tableId), "ID prefix");
  }
  private _idPrefixesByTableId(): Map<string, string> {
    const prefixesInUse = new Set<string>();
    const assigned = new Map<string, string>();
    this._apiAccessTables().forEach((table) => {
      const sampled = table.meta.activeIdPrefix();
      if (sampled === undefined) return;
      prefixesInUse.add(sampled);
      assigned.set(table.tableId, sampled);
    });
    this._apiAccessTables().forEach((table) => {
      if (assigned.has(table.tableId)) return;
      const generated = idPrefixes.fromTitle(table.name, prefixesInUse);
      prefixesInUse.add(generated);
      assigned.set(table.tableId, generated);
    });
    return assigned;
  }
  idPrefixChangeReport(): string | undefined {
    const changes: string[] = [];
    this.tableIdsApiAccesses().forEach((tableId) => {
      const previous = tableConfigsByTableId().get(tableId);
      if (previous === undefined) return;
      const table = this.ss.raw.table(tableId);
      const sampled = table.meta.activeIdPrefix();
      if (sampled === undefined || sampled === previous.idPrefix) return;
      changes.push(
        `Table "${table.name}" sampled ID prefix "${sampled}" differs from last generated "${previous.idPrefix}".`,
      );
    });
    if (changes.length === 0) return undefined;
    return changes.join(" ");
  }
  declaredCellReport(): string | undefined {
    const lines = this.tableConfigSync.declaredCellReportLines;
    if (lines.length === 0) return undefined;
    return lines.join(" ");
  }
  newTableConfigs(): TableConfigsBase {
    const tableConfigs: TableConfigsBase = {};
    const idPrefixLabels: IdPrefixLabel[] = [];
    this._apiAccessTables().forEach((table) => {
      const { tableId, name: tableName } = table;
      const tableKey = this.schema.titleToName(tableName);
      const existing = tableConfigs[tableKey];
      if (existing !== undefined) {
        throw new Error(
          `Tables "${existing.tableName}" and "${tableName}" both give the key "${tableKey}".`,
        );
      }
      const idPrefix = this.idPrefix(tableId);
      const { headerRowIndex, startColIndex } = table.origin;
      const { tableHeaderRow } = table.meta;
      tableConfigs[tableKey] = {
        tableId,
        tableName,
        sheetGid: table.sheetGid,
        idPrefix,
        headerRowIndex,
        startColIndex,
        hasIdColumn: tableHeaderRow.hasValue(this.schema.idHeader),
        hasNameColumn: tableHeaderRow.hasValue(this.schema.nameHeader),
      };
      idPrefixLabels.push({ label: tableName, idPrefix });
    });
    idPrefixes.assertUnique(idPrefixLabels);
    return tableConfigs;
  }
  private _apiAccessTables(): TableRaw[] {
    return this.tableIdsApiAccesses().map((tableId) =>
      this.ss.raw.table(tableId),
    );
  }
  parseColumnReference(reference: string): ColumnReference {
    const match = reference.match(/^(.+)\[(.+)\]$/);
    if (match === null) {
      throw new Error(
        `${columnReferenceLabel(reference)} is not of the form Table[Header].`,
      );
    }
    const tableName = Val.assert(match[1], "Table name match");
    const header = Val.assert(match[2], "header match");
    const table = this._managedTable(tableName, reference);
    const columnId = table.meta.columnIdByHeader(header);
    if (columnId === "") {
      throw new Error(
        `${columnReferenceLabel(reference)} names a column with no column ID.`,
      );
    }
    return { tableId: table.tableId, columnId };
  }
  private _managedTable(tableName: string, reference: string): TableRaw {
    for (const tableId of this.tableIdsApiAccesses()) {
      const table = this.ss.raw.table(tableId);
      if (table.name === tableName) return table;
    }
    throw new Error(
      `${columnReferenceLabel(reference)} names no managed Table "${tableName}".`,
    );
  }
  tableKeysByTableId(): Map<string, string> {
    const map = new Map<string, string>();
    Object.entries(this.newTableConfigs()).forEach(([tableKey, config]) => {
      map.set(config.tableId, tableKey);
    });
    return map;
  }
  toFileSource(makeConfigsImport: string): string {
    return [
      makeImportLine("makeTableConfigs", makeConfigsImport),
      ``,
      `export const tableConfigs = makeTableConfigs(${oneLinePerEntryFileSource(
        this.newTableConfigs(),
      )});`,
      ``,
    ].join("\n");
  }
}

function columnReferenceLabel(reference: string): string {
  return `Column reference "${reference}"`;
}
