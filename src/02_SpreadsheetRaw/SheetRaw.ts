import { Val } from "@byronbroughten/utils/val";

import type {
  ConditionalFormatRule,
  ModelableConditionalFormatRule,
} from "../00_Source/RawSource/ConditionalFormat";
import type {
  EditProtection,
  EditProtectionContent,
  ProtectionGridRange,
  WholeSheetEditLockDeclaration,
  WholeSheetEditWarningDeclaration,
} from "../00_Source/RawSource/EditProtection";
import type {
  GridRangeProps,
  SheetSnapshot,
  TableSnapshot,
} from "../00_Source/RawSource/RawSource";
import {
  type SheetColIndex,
  SheetIndex,
  type SheetRowIndex,
} from "../00_Source/RawSource/SheetIndex";
import { headerZone } from "../01_SpreadsheetSchema/headerZone";
import {
  SpreadsheetBaseRaw,
  type SpreadsheetRawProps,
} from "./ClassBases/SpreadsheetBaseRaw";
import {
  hasQueuedWrites,
  integrateColumnProperties,
  sheetLabel,
  tableIdsOnSheet,
  tableStateOf,
} from "./ClassBases/TableBaseRaw";
import { rowIndexesStaleMessage } from "./ClassBases/TableCommonRaw";
import { emptyStateRaw } from "./ClassTypes/emptyStateRaw";
import type { SheetStateRaw, TableStateRaw } from "./ClassTypes/StateRaw";
import { SheetConditionalFormatsRaw } from "./SheetRaw/SheetConditionalFormatsRaw";
import { SheetEditProtectionsRaw } from "./SheetRaw/SheetEditProtectionsRaw";
import { TableRaw } from "./TableRaw";

export interface SheetRawProps extends SpreadsheetRawProps {
  sheetGid: number;
}

export class SheetRaw extends SpreadsheetBaseRaw {
  readonly sheetGid: number;
  constructor({ sheetGid, ...props }: SheetRawProps) {
    super(props);
    this.sheetGid = sheetGid;
    this._ensureSheetState();
  }
  private _ensureSheetState(): void {
    if (this.sheetsStateRaw.has(this.sheetGid)) return;
    this.sheetsStateRaw.set(this.sheetGid, emptyStateRaw.sheetState());
  }
  private get conditionalFormats(): SheetConditionalFormatsRaw {
    return new SheetConditionalFormatsRaw(this.sheetRawProps);
  }
  private get protections(): SheetEditProtectionsRaw {
    return new SheetEditProtectionsRaw(this.sheetRawProps);
  }
  get sheetRawProps(): SheetRawProps {
    return { ...this.spreadsheetRawProps, sheetGid: this.sheetGid };
  }
  get sheetState(): SheetStateRaw {
    return Val.assert(
      this.sheetsStateRaw.get(this.sheetGid),
      `sheetState for sheetGid ${this.sheetGid}`,
    );
  }
  get hasGatheredFetch(): boolean {
    const { fetchQueue } = this.sheetState;
    return (
      fetchQueue.gatherConditionalFormats || fetchQueue.gatherEditProtections
    );
  }
  // Brings every Table on the sheet without counting any of them as used.
  gatherFetchHeaderZone(): this {
    const { fetchQueue } = this.sheetState;
    if (fetchQueue.gatherHeaderZone) return this;
    fetchQueue.gatherHeaderZone = true;
    this.spreadsheetStateRaw.fetchQueue.gridRanges.push({
      sheetId: this.sheetGid,
      startRowIndex: SheetIndex.row(0),
      endRowIndex: headerZone.endRowIndex,
    });
    return this;
  }
  get title(): string {
    const title = this.sheetState.working.title;
    if (title === undefined) {
      throw new Error(
        `Sheet title is null for sheetGid ${this.sheetGid}. Ensure that the sheet properties have been fetched.`,
      );
    }
    return title;
  }
  get label(): string {
    return sheetLabel(this.sheetState.working.title, this.sheetGid);
  }
  updateTitle(title: string): this {
    this.writeOperations.renameSheet.push({
      kind: "renameSheet",
      sheetId: this.sheetGid,
      title,
    });
    this.sheetState.working.title = title;
    return this;
  }
  get rowCount(): number {
    return Val.assert(
      this.sheetState.working.rowCount,
      `${this.label}'s row count`,
    );
  }
  get columnCount(): number {
    return Val.assert(
      this.sheetState.working.columnCount,
      `${this.label}'s column count`,
    );
  }
  queueGridRowsThrough(endRowIndex: SheetRowIndex): void {
    const { rowCount } = this;
    if (endRowIndex <= rowCount) return;
    this.sheetState.writeQueue.appendedRowCount += endRowIndex - rowCount;
    this.sheetState.working.rowCount = endRowIndex;
  }
  queueGridColumnsThrough(endColumnIndex: SheetColIndex): void {
    const { columnCount } = this;
    if (endColumnIndex <= columnCount) return;
    this.sheetState.writeQueue.appendedColumnCount +=
      endColumnIndex - columnCount;
    this.sheetState.working.columnCount = endColumnIndex;
  }
  get wholeSheetGridRange(): ProtectionGridRange {
    return { sheetId: this.sheetGid };
  }
  // A range with row coordinates can cover any of the sheet's Tables, so none may have moved rows.
  assertRowIndexesNotStale(): void {
    const stale = this.tableStates.find(
      (tableState) => tableState.properties?.rowIndexesAreStale === true,
    );
    if (stale?.properties === undefined) return;
    throw new Error(
      rowIndexesStaleMessage(
        `Table "${stale.properties.name}" on ${this.label}`,
      ),
    );
  }
  private get tableStates(): TableStateRaw[] {
    const onSheet = Array.from(this.tablesStateRaw.values()).filter(
      (tableState) => tableState.sheetGid === this.sheetGid,
    );
    return [...onSheet, ...this.sheetState.tablesBeforePropertiesById.values()];
  }
  gatherFetchConditionalFormatRules(): this {
    this.conditionalFormats.gatherFetchConditionalFormatRules();
    return this;
  }
  conditionalFormatRules(): ConditionalFormatRule[] {
    return this.conditionalFormats.conditionalFormatRules();
  }
  hasPendingConditionalFormatRule(
    rule: ModelableConditionalFormatRule,
    queuedOnTable: ModelableConditionalFormatRule[],
  ): boolean {
    return this.conditionalFormats.hasPendingConditionalFormatRule(
      rule,
      queuedOnTable,
    );
  }
  removeConditionalFormatRulesAt(range: GridRangeProps): this {
    this.conditionalFormats.removeConditionalFormatRulesAt(range);
    return this;
  }
  removeConditionalFormatRule(rule: ConditionalFormatRule): this {
    this.conditionalFormats.removeConditionalFormatRule(rule);
    return this;
  }
  markConditionalFormatIndexesStale(): void {
    this.conditionalFormats.markConditionalFormatIndexesStale();
  }
  assertConditionalFormatIndexesNotStale(): void {
    this.conditionalFormats.assertConditionalFormatIndexesNotStale();
  }
  integrateConditionalFormatRules(rules: ConditionalFormatRule[]): void {
    this.conditionalFormats.integrateConditionalFormatRules(rules);
  }
  gatherFetchEditProtections(): this {
    this.protections.gatherFetchEditProtections();
    return this;
  }
  editProtections(): EditProtection[] {
    return this.protections.editProtections();
  }
  addEditWarningWholeSheet(
    declaration: WholeSheetEditWarningDeclaration = {},
  ): this {
    this.protections.addEditWarningWholeSheet(declaration);
    return this;
  }
  addEditLockWholeSheet(declaration: WholeSheetEditLockDeclaration = {}): this {
    this.protections.addEditLockWholeSheet(declaration);
    return this;
  }
  hasPendingEditProtection(
    protection: EditProtectionContent,
    queuedOnTable: EditProtectionContent[],
  ): boolean {
    return this.protections.hasPendingEditProtection(protection, queuedOnTable);
  }
  removeEditProtectionsAt(range: ProtectionGridRange): this {
    this.protections.removeEditProtectionsAt(range);
    return this;
  }
  removeEditProtection(protection: EditProtection): this {
    this.protections.removeEditProtection(protection);
    return this;
  }
  removeEditProtectionByDescription(description: string): this {
    this.protections.removeEditProtectionByDescription(description);
    return this;
  }
  removeEditProtectionById(protectionId: number): this {
    this.protections.removeEditProtectionById(protectionId);
    return this;
  }
  markEditProtectionsStale(): void {
    this.protections.markEditProtectionsStale();
  }
  assertEditProtectionsNotStale(): void {
    this.protections.assertEditProtectionsNotStale();
  }
  integrateEditProtections(protections: EditProtection[]): void {
    this.protections.integrateEditProtections(protections);
  }
  // A GID reaches the sheet, and the sheet reaches a Table by its ID, before its properties arrive too.
  table(tableId: string): TableRaw {
    return new TableRaw({
      ...this.spreadsheetRawProps,
      tableId,
      sheetGid: this.sheetGid,
    });
  }
  // Every Table fetched onto the sheet, managed or not.
  get tableIds(): string[] {
    return tableIdsOnSheet(this.tablesStateRaw, this.sheetGid);
  }
  // Row indexes only actually shift once the deletes have been sent, and a Table below the deleted rows shifts with them.
  markRowIndexesStale(): void {
    this.tableIds.forEach((tableId) =>
      this.table(tableId).markRowIndexesStale(),
    );
  }
  integrateSnapshot(sheet: SheetSnapshot): void {
    this._integrateSheetProperties(sheet);
    this.tableIds.forEach((tableId) => {
      this.table(tableId).integrateGridBlocks(sheet.gridBlocks ?? []);
    });
  }
  removeTablesAbsentFrom(tables: TableSnapshot[]): void {
    const liveTableIds = tables.map(({ tableId }) => tableId);
    this.tableIds
      .filter((tableId) => !liveTableIds.includes(tableId))
      .forEach((tableId) => this._removeAbsentTable(tableId));
  }
  // A queued write must never vanish with its Table, so it stops the run instead.
  private _removeAbsentTable(tableId: string): void {
    const { writeQueue } = tableStateOf(this.tablesStateRaw, tableId);
    if (hasQueuedWrites(writeQueue)) {
      throw new Error(
        `Table ${tableId} is no longer on ${this.label}, but it has queued writes; refetch before queuing writes to it.`,
      );
    }
    this.tablesStateRaw.delete(tableId);
  }
  private _integrateSheetProperties(sheet: SheetSnapshot): void {
    if (sheet.title) {
      this.sheetState.working.title = sheet.title;
    }
    if (sheet.rowCount !== undefined) {
      this.sheetState.working.rowCount = sheet.rowCount;
    }
    if (sheet.columnCount !== undefined) {
      this.sheetState.working.columnCount = sheet.columnCount;
    }
    if (sheet.tables !== undefined) {
      this._integrateTables(sheet.tables);
    }
    this._integrateQueuedSheetProperties();
  }
  // A grid fetch returns only the Tables its ranges overlap, so a Table it leaves out stays.
  private _integrateTables(tables: TableSnapshot[]): void {
    tables.forEach((table) => {
      const tableState =
        this.tablesStateRaw.get(table.tableId) ??
        this._takeTableBeforePropertiesById(table.tableId) ??
        emptyStateRaw.tableState(this.sheetGid);
      tableState.sheetGid = this.sheetGid;
      tableState.properties = {
        tableId: table.tableId,
        name: table.name,
        startRowIndex: table.startRowIndex,
        endRowIndex: table.endRowIndex,
        startColumnIndex: table.startColumnIndex,
        endColumnIndex: table.endColumnIndex,
        columnProperties: table.columnProperties,
        rowIndexesAreStale: tableState.properties?.rowIndexesAreStale ?? false,
      };
      integrateColumnProperties(tableState, table);
      this.tablesStateRaw.set(table.tableId, tableState);
    });
  }
  private _takeTableBeforePropertiesById(
    tableId: string,
  ): TableStateRaw | undefined {
    const { tablesBeforePropertiesById } = this.sheetState;
    const tableState = tablesBeforePropertiesById.get(tableId);
    tablesBeforePropertiesById.delete(tableId);
    return tableState;
  }
  // Queued properties outlive a re-fetch until the flush sends them.
  private _integrateQueuedSheetProperties(): void {
    this.writeOperations.renameSheet.forEach(({ sheetId, title }) => {
      if (sheetId === this.sheetGid) this.sheetState.working.title = title;
    });
    this.writeOperations.renameTable.forEach(({ tableId, name }) => {
      const properties = this.tablesStateRaw.get(tableId)?.properties;
      if (properties !== undefined) properties.name = name;
    });
  }
}
