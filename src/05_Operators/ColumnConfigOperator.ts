import { getColumnTraitByName } from "../01_SpreadsheetSchema/columnConfigsTypes";
import {
  configSheetFloorSeed,
  floorSeedColumnById,
} from "../01_SpreadsheetSchema/configSheetFloorSeed";
import {
  type ColumnConfigsGeneric,
  makeImportLine,
} from "../01_SpreadsheetSchema/makeConfigs";
import { type ValueName } from "../01_SpreadsheetSchema/valueSchemas";
import type { ColumnMetaRaw } from "../02_SpreadsheetRaw/ColumnMetaRaw";
import { Str } from "../utils/Str";
import { columnConfigsFileSource } from "./configFileSource";
import { GenericSheetOperator } from "./GenericSheetOperator";
import { SheetConfigOperator } from "./SheetConfigOperator";
import {
  type ConfigSyncState,
  type OperatorProps,
  SpreadsheetBaseOperator,
  type UntypedHeadersBySheetTitle,
} from "./SpreadsheetBaseOperator";
import { ValueConfigOperator } from "./ValueConfigOperator";

interface ColumnIdentity {
  sheetGid: number;
  columnId: string;
}

export class ColumnConfigOperator extends GenericSheetOperator<"columnConfig"> {
  constructor(props: OperatorProps) {
    super({
      sheetName: "columnConfig",
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
  get sheetConfigOperator(): SheetConfigOperator {
    return new SheetConfigOperator(this.operatorProps);
  }
  get valueConfigOperator(): ValueConfigOperator {
    return new ValueConfigOperator(this.operatorProps);
  }
  private get untypedHeadersBySheetTitle(): UntypedHeadersBySheetTitle {
    return this.columnConfigSync.untypedHeadersBySheetTitle;
  }
  // Derived fresh each call, not cached — a stored field goes stale across
  // this coordinator's per-access getter rebuilds.
  private get sheetGidsApiAccesses(): Set<number> {
    return new Set(this.sheetConfigOperator.sheetGidsApiAccesses());
  }
  activeValueTitles(): string[] {
    return this.sheet.rowIndexesActiveWithData.map((rowIndex) =>
      this._describedColumn(this._columnIdentity(rowIndex)).activeValueTitle(),
    );
  }
  assertSyncedToSpreadsheet(): void {
    if (!this.columnConfigSync.syncedToSpreadsheet) {
      throw new Error(
        "ColumnConfigOperator has not yet synced to the spreadsheet.",
      );
    }
  }
  prepFetchWithSheetConfig(): void {
    this.sheetConfigOperator.assertPrepFetchIsComplete();
    this.sheet.prepFetchColumnsFull(
      "sheetGid",
      "columnId",
      "sheetTitle",
      "header",
      "emptyValueAllowed",
    );
  }
  fetchAfterSheetConfigSynced(): this {
    this.sheetConfigOperator.assertSyncedToSpreadsheet();
    this.sheetGidsApiAccesses.forEach((sheetGid) => {
      const sheet = this.ss.raw.sheet(sheetGid);
      // hasIdColumn samples this row after Let api access is known.
      sheet.meta.tableHeaderRow.gatherFetchFull();
      sheet.meta.colIdRow.gatherFetchFull();
      sheet.topRow.gatherFetchFull();
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
    const untypedHeaders = Array.from(this.untypedHeadersBySheetTitle.values());
    if (untypedHeaders.length === 0) {
      return undefined;
    }
    const columnCount = untypedHeaders.reduce(
      (count, headers) => count + headers.length,
      0,
    );
    const sentences = [
      `Succeeded, but ${columnCount} column(s) across ${untypedHeaders.length} ` +
        `sheet(s) are untyped, so their value names were guessed. See the ` +
        `execution log for the list.`,
    ];
    const blankSampleTitles = this._blankSampleSheetTitles();
    if (blankSampleTitles.length > 0) {
      const names = blankSampleTitles.map((title) => `"${title}"`).join(", ");
      sentences.push(
        `On ${blankSampleTitles.length} of those sheet(s) the top data row ` +
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
  private _blankSampleSheetTitles(): string[] {
    const titles: string[] = [];
    this.sheetGidsApiAccesses.forEach((sheetGid) => {
      const sheet = this.ss.raw.sheet(sheetGid);
      if (!this.untypedHeadersBySheetTitle.has(sheet.title)) return;
      if (!sheet.topDataRowIsBlank()) return;
      titles.push(sheet.title);
    });
    return titles;
  }
  private _isSheetGidApiAccesses(sheetGid: number): boolean {
    return this.sheetGidsApiAccesses.has(sheetGid);
  }
  private _addMissingColumnIds(): this {
    let idsAdded = 0;
    this.sheetGidsApiAccesses.forEach((sheetGid) => {
      const idPrefix = this.sheetConfigOperator.idPrefix(sheetGid);
      idsAdded += this.ss.raw.sheetMeta(sheetGid).addMissingColumnIds(idPrefix);
    });
    Logger.log(
      `ensureColumnIds: prepared to add ${idsAdded} missing column ID(s)`,
    );
    return this;
  }
  private _pruneColumnRows(): this {
    const col = this.sheet.columns("sheetGid", "columnId");
    let staleCount = 0;
    this.sheet.rowIndexesActive.forEach((rowIndex) => {
      const sheetGid = col.sheetGid.valueOrEmpty(rowIndex);
      const columnId = col.columnId.valueOrEmpty(rowIndex);
      if (
        sheetGid === "" ||
        columnId === "" ||
        !this._isSheetGidApiAccesses(sheetGid) ||
        !this._isActiveColumnId(sheetGid, columnId)
      ) {
        this.sheet.row(rowIndex).delete();
        staleCount++;
      }
    });
    Logger.log(`_pruneColumnRows: pruned ${staleCount} stale row(s).`);
    return this;
  }
  private _isActiveColumnId(sheetGid: number, columnId: string): boolean {
    return this.ss.raw.sheetMeta(sheetGid).isActiveColumnId(columnId);
  }
  private _appendColumnRows(): this {
    const existingIdentityKeys = new Set(
      this.sheet.rowIndexesActiveWithData.map((rowIndex) =>
        columnIdentityKey(this._columnIdentity(rowIndex)),
      ),
    );

    let appendedCount = 0;
    this.sheetGidsApiAccesses.forEach((sheetGid) => {
      const { activeColumnIds } = this.ss.raw.sheetMeta(sheetGid);
      activeColumnIds.forEach((columnId) => {
        if (
          !existingIdentityKeys.has(columnIdentityKey({ sheetGid, columnId }))
        ) {
          this.sheet.appendRowWithVals({ sheetGid, columnId });
          appendedCount++;
        }
      });
    });
    Logger.log(`_appendColumnRows: added ${appendedCount} new row(s).`);
    return this;
  }
  private _columnIdentity(rowIndex: number): ColumnIdentity {
    const col = this.sheet.columns("sheetGid", "columnId");
    return {
      sheetGid: col.sheetGid.value(rowIndex),
      columnId: col.columnId.value(rowIndex),
    };
  }
  private _describedColumn({
    sheetGid,
    columnId,
  }: ColumnIdentity): ColumnMetaRaw {
    return this.ss.raw.sheetMeta(sheetGid).columnByActiveId(columnId);
  }
  private _updateProgrammaticValues(): void {
    const col = this.sheet.columns("sheetTitle", "header", "emptyValueAllowed");
    const reportLines = this.columnConfigSync.declaredCellReportLines;
    reportLines.length = 0;
    let updatedValues = 0;
    this.untypedHeadersBySheetTitle.clear();
    this.sheet.rowIndexesActiveWithData.forEach((rowIndex) => {
      const identity = this._columnIdentity(rowIndex);
      const sheetRaw = this.ss.raw.sheet(identity.sheetGid);

      const actualSheetTitle = sheetRaw.title;
      if (col.sheetTitle.valueOrEmpty(rowIndex) !== actualSheetTitle) {
        col.sheetTitle.cell(rowIndex).updateValue(actualSheetTitle);
        updatedValues++;
      }

      const describedColumn = this._describedColumn(identity);
      const actualHeader = describedColumn.activeHeader;
      if (col.header.valueOrEmpty(rowIndex) !== actualHeader) {
        col.header.cell(rowIndex).updateValue(actualHeader);
        updatedValues++;
      }

      if (
        this._updateSelfDescribingEmptyValueAllowed({
          rowIndex,
          identity,
          sheetTitle: actualSheetTitle,
          header: actualHeader,
        })
      ) {
        updatedValues++;
      }

      if (describedColumn.activeDeclaredValueTitle() === undefined) {
        this._recordUntypedColumn(actualSheetTitle, actualHeader);
      }
    });
    Logger.log(`Corrected ${updatedValues} inaccurate Column Config cell(s).`);
  }
  private _updateSelfDescribingEmptyValueAllowed({
    rowIndex,
    identity,
    sheetTitle,
    header,
  }: {
    rowIndex: number;
    identity: ColumnIdentity;
    sheetTitle: string;
    header: string;
  }): boolean {
    const seedColumn = floorSeedColumnById(
      identity.sheetGid,
      identity.columnId,
    );
    if (seedColumn === undefined) return false;
    const emptyValueAllowed = this.sheet.column("emptyValueAllowed");
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
      )} · ${sheetTitle} · ${header} → ${
        seedColumn.emptyValueAllowed ? "TRUE" : "FALSE"
      }`,
    );
    return true;
  }
  private _recordUntypedColumn(sheetTitle: string, header: string): void {
    const headers = this.untypedHeadersBySheetTitle.get(sheetTitle) ?? [];
    headers.push(header);
    this.untypedHeadersBySheetTitle.set(sheetTitle, headers);
  }
  private _logUntypedColumns(): void {
    this.untypedHeadersBySheetTitle.forEach((headers, sheetTitle) => {
      Logger.log(
        `Untyped columns on "${sheetTitle}" (${headers.length}): ${headers.join(", ")}`,
      );
    });
  }
  newColumnConfigs(): ColumnConfigsGeneric {
    const sheetNamesByGid = this.sheetConfigOperator.sheetNamesByGid();
    const col = this.sheet.columns("header", "emptyValueAllowed");
    const columnConfigs: ColumnConfigsGeneric = {};
    this.sheet.rowIndexesActiveWithData.forEach((rowIndex) => {
      const identity = this._columnIdentity(rowIndex);
      const { sheetGid, columnId } = identity;
      const header = col.header.value(rowIndex);
      const sheetName = sheetNamesByGid.get(sheetGid);
      if (!sheetName) {
        throw new Error(
          `generateColumnConfigFileSource: column "${columnId}" references sheetGid ` +
            `${sheetGid}, which has no corresponding sheet name in Sheet Config.`,
        );
      }
      const columnName = Str.sentenceToCamelCase(header);
      if (!columnConfigs[sheetName]) {
        columnConfigs[sheetName] = {};
      }
      const tableColumnConfigs = columnConfigs[sheetName];
      if (tableColumnConfigs[columnName]) {
        throw new Error(
          `generateColumnConfigFileSource: duplicate column name "${columnName}" ` +
            `derived from header "${header}" on sheet "${sheetName}".`,
        );
      }
      const describedColumn = this._describedColumn(identity);
      tableColumnConfigs[columnName] = {
        columnId,
        header,
        valueName: this.schema.titleToName(
          describedColumn.activeValueTitle(),
        ) as ValueName,
        isFormula: describedColumn.activeIsFormula,
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

function columnIdentityKey({ sheetGid, columnId }: ColumnIdentity): string {
  return `${sheetGid}:${columnId}`;
}
