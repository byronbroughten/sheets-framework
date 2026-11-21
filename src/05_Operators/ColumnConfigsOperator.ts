import { Val } from "@byronbroughten/utils/val";

import {
  type BraceConfigs,
  BraceFormula,
  type BraceNames,
} from "../01_SpreadsheetSchema/configReaders/BraceFormula";
import { columnNameOfHeader } from "../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import type { ColumnConfigsGeneric } from "../01_SpreadsheetSchema/configReaders/makeConfigs";
import type { FormulaColumn } from "../01_SpreadsheetSchema/configReaders/OutsideReferences";
import { previousTraitTables } from "../01_SpreadsheetSchema/configReaders/previousTraitTables";
import {
  type IntendedTrait,
  newColumnLines,
  newTableLines,
  type PlannedColumnGroupMove,
  type PlannedNewColumn,
  type PlannedNewTable,
  type PlanReadings,
  type TraitReadings,
} from "../01_SpreadsheetSchema/configReaders/SyncPlanner";
import type { WideColumnGroupMove } from "../01_SpreadsheetSchema/configReaders/SyncPlanner/wideAlterations";
import {
  beforeFirstColumnText,
  type SyncBeforeFirstColumnStop,
  type SyncCreatedColumn,
  type SyncCreatedTable,
  type SyncHeaderTakenStop,
  type SyncSendLine,
} from "../01_SpreadsheetSchema/configReaders/syncRunResult";
import { dimensionIds } from "../01_SpreadsheetSchema/dimensionIds";
import type {
  ColumnMoveSpan,
  NewColumnHeadCells,
} from "../02_SpreadsheetRaw/ClassTypes/StateRaw";
import { ColumnConfigsBaseOperator } from "./ClassBases/ColumnConfigsBaseOperator";
import type { NamedColumnIdentity } from "./ColumnConfigsOperator/ApiColumnIdentities";
import {
  type ApiTableColumns,
  ColumnConfigsOutputOperator,
} from "./ColumnConfigsOperator/ColumnConfigsOutputOperator";
import { ColumnValueSourcesOperator } from "./ColumnConfigsOperator/ColumnValueSourcesOperator";
import {
  bracedFormula,
  InstalledBraceConfigs,
} from "./ColumnConfigsOperator/InstalledBraceConfigs";
import { TableColumnOrder } from "./ColumnConfigsOperator/TableColumnOrder";
import { UntypedColumnsOperator } from "./ColumnConfigsOperator/UntypedColumnsOperator";
import {
  SpreadsheetBaseOperator,
  type ValueSource,
} from "./SpreadsheetBaseOperator";
import type { MintedNewTable } from "./TableConfigsOperator";

interface SentColumnTrait extends NamedColumnIdentity {
  value: IntendedTrait["value"];
}

export interface MintedNewColumn extends PlannedNewColumn {
  tableId: string;
  columnId: string;
}

export interface LandedColumns {
  createdColumns: SyncCreatedColumn[];
  didntLand: SyncSendLine[];
}

export interface LandedTables {
  createdTables: SyncCreatedTable[];
  didntLand: SyncSendLine[];
}

interface ColumnAt {
  tableId: string;
  colIndex: number;
}

interface ColumnInsert extends ColumnAt {
  headCells: NewColumnHeadCells;
}

// Indexed in the layout the run's column deletions and inserts leave; `table` is the codebase name a rule rewrite looks it up by.
export interface PlannedColumnMove extends ColumnMoveSpan {
  tableId: string;
  table: string;
}

interface ColumnMovesProps {
  sent: IntendedTrait[];
  minted: MintedNewColumn[];
  groupMoves: PlannedColumnGroupMove[];
}

// Columns to sit together, in order, immediately before `before`, all by previous names.
interface OrderedColumns extends WideColumnGroupMove {
  tableId: string;
  tableKey: string;
}

interface ColumnOrderProps {
  tableId: string;
  tableKey: string;
  deleted: Set<string>;
  minted: MintedNewColumn[];
}

export class ColumnConfigsOperator extends ColumnConfigsBaseOperator {
  static init(): ColumnConfigsOperator {
    return new ColumnConfigsOperator(
      SpreadsheetBaseOperator.initOperatorProps(),
    );
  }
  get valueSourcesOperator(): ColumnValueSourcesOperator {
    return new ColumnValueSourcesOperator(this.operatorProps);
  }
  get untypedColumnsOperator(): UntypedColumnsOperator {
    return new UntypedColumnsOperator(this.operatorProps);
  }
  get outputOperator(): ColumnConfigsOutputOperator {
    return new ColumnConfigsOutputOperator(this.operatorProps);
  }
  valueSources(): ValueSource[] {
    return this.valueSourcesOperator.valueSources();
  }
  assertSyncedToSpreadsheet(): void {
    this.untypedColumnsOperator.assertSyncedToSpreadsheet();
  }
  fetchHeadRows(): this {
    return this.fetchHeadRowsOf([
      ...this.apiColumnIdentities.tableIdsApiAccesses,
    ]);
  }
  fetchHeadRowsOf(tableIds: string[]): this {
    for (const tableId of tableIds) {
      const table = this.ss.raw.table(tableId);
      // hasIdColumn samples this row after API access is known.
      table.headRow("header").gatherFetchFull();
      table.headRow("columnId").gatherFetchFull();
      table.headRow("action").gatherFetchFull();
      table.headRow("groupHeading1").gatherFetchFull();
      table.topRow.gatherFetchFull();
    }
    this.ss.identified.fetchAllGathered(true);
    return this;
  }
  addMissingColumnIds(): this {
    let idsAdded = 0;
    const idPrefixesByTableId = this.tableConfigOperator.idPrefixesByTableId();
    for (const tableId of this.apiColumnIdentities.tableIdsApiAccesses) {
      const idPrefix = idPrefixesByTableId.assertIdPrefix(tableId);
      // Raw, not TableIdentified: a Table here may predate its generated schema.
      const table = this.ss.raw.table(tableId);
      const colIndexes = table.columnResolver.colIndexesWithoutColumnId;
      for (const colIndex of colIndexes) {
        table
          .headRow("columnId")
          .updateValue(colIndex, dimensionIds.col(idPrefix));
      }
      idsAdded += colIndexes.length;
    }
    Logger.log(
      `ensureColumnIds: prepared to add ${idsAdded} missing column ID(s)`,
    );
    return this;
  }
  recordUntypedColumns(): this {
    this.untypedColumnsOperator.recordUntypedColumns();
    return this;
  }
  // emptyValueAllowed isn't a sheet trait, so its live side carries the previous configs and only a deleted column can stop it.
  // `before` is read live only, as a tab's is: order is never in the configs, so it can't conflict.
  traitReadings(): PlanReadings {
    const liveConfigs = this._liveBraceConfigs();
    const { apiColumnIdentities } = this;
    const identities = apiColumnIdentities.namedColumnIdentities();
    const namesByColumnId = new Map(
      identities.map(({ columnId, columnName }) => [columnId, columnName]),
    );
    const live = identities.reduce<TraitReadings["tables"]>(
      (tables, identity) => {
        const { tableId, tableKey, columnId, columnName } = identity;
        const reading = tables[tableKey] ?? {
          traits: { name: this.ss.raw.table(tableId).name },
          columns: {},
        };
        const { header, columnType, validationRule } =
          apiColumnIdentities.describedColumn(identity);
        const next = this._columnOrder(tableId, tableKey).columnIdAfter(
          columnId,
        );
        reading.columns[columnName] = {
          header,
          emptyValueAllowed: identity.previousEmptyValueAllowed,
          formula: this._bracedTopFormula(liveConfigs, identity),
          // Against live names under previous keys, so a person renaming a Table a rule names reads as no change.
          columnType: BraceFormula.decompileColumnType({
            configs: liveConfigs,
            tableKey,
            columnType: {
              type: columnType ?? null,
              rule: validationRule ?? null,
            },
          }),
          before: next === undefined ? undefined : namesByColumnId.get(next),
        };
        tables[tableKey] = reading;
        return tables;
      },
      {},
    );
    return {
      previous: {
        tables: previousTraitTables(this.installedBraceConfigs.installed),
        tabs: {},
      },
      live: { tables: live, tabs: {} },
    };
  }
  // The live formula as traitReadings brace it, not the installed configs' text, so a person's edit since the last sync is what's checked.
  formulaColumns(): FormulaColumn[] {
    const configs = this._liveBraceConfigs();
    return this.apiColumnIdentities
      .namedColumnIdentities()
      .flatMap((identity) => {
        const { tableId, tableKey, columnId } = identity;
        const formula = this._bracedTopFormula(configs, identity);
        if (formula === undefined) return [];
        const cell = this.ss.raw
          .table(tableId)
          .columnBodyA1(this._colIndexOf(tableId, columnId));
        return [{ tableKey, cell, formula }];
      });
  }
  // In brace form against the configs its callers pass, live configs for a live formula.
  private _bracedTopFormula(
    configs: BraceConfigs,
    identity: NamedColumnIdentity,
  ): string | undefined {
    const { topFormula } = this.apiColumnIdentities.describedColumn(identity);
    return (
      bracedFormula({
        configs,
        tableKey: identity.tableKey,
        formula: topFormula,
      }) ?? undefined
    );
  }
  private _liveBraceConfigs(): BraceConfigs {
    return this.installedBraceConfigs.liveNamed(
      this.ss.raw,
      this.apiColumnIdentities.tableIdsApiAccesses,
    );
  }
  // Queues each sent column header, addressed by the previous names the readings use.
  updateHeaders(sent: IntendedTrait[]): this {
    for (const { tableId, columnId, value } of this._sentColumnTraits(
      sent,
      "header",
    )) {
      this.ss.raw.table(tableId).updateHeader(columnId, String(value));
    }
    return this;
  }
  // Fills every body row with one formula; a Table copies a uniform column formula into the rows it grows by.
  updateFormulas(sent: IntendedTrait[]): this {
    const configs = this.installedBraceConfigs.atColumnFill(sent);
    for (const { tableId, columnId, tableKey, value } of this._sentColumnTraits(
      sent,
      "formula",
    )) {
      const table = this.ss.raw.table(tableId);
      table.column(table.columnResolver.colIndexOf(columnId)).updateAllFormulas(
        BraceFormula.compileForCell({
          configs,
          tableKey,
          formula: String(value),
        }),
      );
    }
    return this;
  }
  // Each with its own rule, so a dropdown's values change in the same send as its type.
  updateColumnTypes(sent: IntendedTrait[]): this {
    const configs = this.installedBraceConfigs.atColumnFill(sent);
    for (const { tableId, columnId, tableKey, value } of this._sentColumnTraits(
      sent,
      "columnType",
    )) {
      if (typeof value !== "object") {
        throw new Error(
          `The columnType sent to ${tableKey} column ${columnId} isn't a type and rule.`,
        );
      }
      const table = this.ss.raw.table(tableId);
      table
        .column(table.columnResolver.colIndexOf(columnId))
        .updateColumnTypeAndRule(
          BraceFormula.compileColumnType({
            configs,
            tableKey,
            columnType: value,
          }),
        );
    }
    return this;
  }
  // Matched by the name the regen would give the header, since two headers giving one name can't both be configured.
  headerTakenStops(newColumns: PlannedNewColumn[]): SyncHeaderTakenStop[] {
    const { apiColumnIdentities } = this;
    const identities = apiColumnIdentities.namedColumnIdentities();
    return newColumns.flatMap(({ table, column, header }) => {
      const taken = identities.find(
        (identity) =>
          identity.tableKey === table &&
          columnNameOfHeader(
            apiColumnIdentities.describedColumn(identity).header,
          ) === column,
      );
      if (taken === undefined) return [];
      const { columnName, columnId } = taken;
      return [
        { kind: "headerTaken", table, header, column: columnName, columnId },
      ];
    });
  }
  // From the live column order, which the previous ID-name map doesn't hold, so the entry refuser can't catch it before the read.
  beforeFirstColumnStops(
    newColumns: PlannedNewColumn[],
  ): SyncBeforeFirstColumnStop[] {
    return newColumns.flatMap(
      ({ table, beforePath, header, before }): SyncBeforeFirstColumnStop[] => {
        if (before === undefined) return [];
        const placedBefore = this._liveColumnNamed(table, before);
        if (placedBefore === undefined) return [];
        if (!this._isFirstColumn(placedBefore)) return [];
        return [
          {
            kind: "beforeFirstColumn",
            table,
            path: beforePath,
            header,
            before,
          },
        ];
      },
    );
  }
  // Minted before any queue, so a dry run's withheld flush and the report carry the same IDs.
  mintColumnIds(newColumns: PlannedNewColumn[]): MintedNewColumn[] {
    const idPrefixesByTableId = this.tableConfigOperator.idPrefixesByTableId();
    return newColumns.map((newColumn) => {
      const tableId = this.tableConfigOperator.tableIdOf(newColumn.table);
      this._validateInsertable(tableId, newColumn);
      const idPrefix = idPrefixesByTableId.assertIdPrefix(tableId);
      return { ...newColumn, tableId, columnId: dimensionIds.col(idPrefix) };
    });
  }
  // A guard only: beforeFirstColumnStops stops the run first. Before anything is queued, so the throw sends nothing.
  private _validateInsertable(
    tableId: string,
    { table, header, before }: PlannedNewColumn,
  ): void {
    if (before === undefined) return;
    const columnId = this._columnIdOf(table, before);
    if (!this._isFirstColumn({ tableId, columnId })) return;
    throw new Error(beforeFirstColumnText({ table, header, before }));
  }
  // At the end, a new column shifts none of the plan's other writes, so it joins their flush.
  appendColumns(minted: MintedNewColumn[]): this {
    for (const { tableId, columnId, header, before } of minted) {
      if (before !== undefined) continue;
      this.ss.raw.table(tableId).appendColumn({ columnId, header });
    }
    return this;
  }
  // Found by minted ID in the fresh configs, which name each column by its live header.
  landedColumns(minted: MintedNewColumn[]): LandedColumns {
    const columnConfigs = this.newColumnConfigs();
    const landed: LandedColumns = { createdColumns: [], didntLand: [] };
    for (const newColumn of minted) {
      const found = Object.entries(columnConfigs[newColumn.table] ?? {}).find(
        ([, config]) => config.columnId === newColumn.columnId,
      );
      if (found === undefined) {
        landed.didntLand.push(newColumnLines.sendLine(newColumn));
        continue;
      }
      const [column, config] = found;
      landed.createdColumns.push({
        kind: "createdColumn",
        table: newColumn.table,
        header: config.header,
        column,
        columnId: newColumn.columnId,
      });
    }
    return landed;
  }
  // Found by minted ID in the fresh configs, the Table with every one of its columns, or it didn't land.
  landedTables(minted: MintedNewTable[]): LandedTables {
    const columnConfigs = this.newColumnConfigs();
    const tableKeysByTableId = this.tableConfigOperator.tableKeysByTableId();
    const landed: LandedTables = { createdTables: [], didntLand: [] };
    for (const newTable of minted) {
      const tableKey = tableKeysByTableId.get(newTable.tableId);
      const configs = Object.entries(
        tableKey === undefined ? {} : (columnConfigs[tableKey] ?? {}),
      );
      const columns = newTable.columns.flatMap(({ columnId }) => {
        const found = configs.find(
          ([, config]) => config.columnId === columnId,
        );
        if (found === undefined) return [];
        const [column, config] = found;
        return [{ header: config.header, column, columnId }];
      });
      if (tableKey === undefined || columns.length < newTable.columns.length) {
        landed.didntLand.push(newTableLines.sendLine(newTable));
        continue;
      }
      landed.createdTables.push({
        kind: "createdTable",
        table: tableKey,
        tab: this.ss.raw.sheet(newTable.createdPlacement.sheetGid).title,
        tableId: newTable.tableId,
        columns,
      });
    }
    return landed;
  }
  // Rightmost first on each Table, so a delete's index isn't moved by the ones queued before it.
  // Then the inserts, each at its index once those deletes land.
  updateDeletionsAndInserts(
    sent: IntendedTrait[],
    minted: MintedNewColumn[],
  ): this {
    const deletions = this._sentColumnTraits(sent, "deleted").map(
      ({ tableId, columnId }) => ({
        tableId,
        colIndex: this._colIndexOf(tableId, columnId),
      }),
    );
    deletions.sort((first, second) => second.colIndex - first.colIndex);
    for (const { tableId, colIndex } of deletions) {
      this.ss.raw.table(tableId).deleteColumn(colIndex);
    }
    for (const { tableId, colIndex, headCells } of this._inserts(
      minted,
      deletions,
    )) {
      this.ss.raw.table(tableId).insertColumn(colIndex, headCells);
    }
    return this;
  }
  // Rightmost first, and two before one column in reverse entry order, so the entries land in the order written.
  private _inserts(
    minted: MintedNewColumn[],
    deletions: ColumnAt[],
  ): ColumnInsert[] {
    const inserts = minted.flatMap(
      ({ tableId, table, columnId, header, before }): ColumnInsert[] => {
        if (before === undefined) return [];
        const beforeIndex = this._colIndexOf(
          tableId,
          this._columnIdOf(table, before),
        );
        const deletedLeft = deletions.filter(
          (deletion) =>
            deletion.tableId === tableId && deletion.colIndex < beforeIndex,
        ).length;
        return [
          {
            tableId,
            colIndex: beforeIndex - deletedLeft,
            headCells: { columnId, header },
          },
        ];
      },
    );
    return inserts
      .reverse()
      .sort((first, second) => second.colIndex - first.colIndex);
  }
  // Before any queue, so a throw sends nothing; each move is indexed in the layout the run's deletions, inserts and earlier moves leave.
  planColumnMoves({
    sent,
    minted,
    groupMoves,
  }: ColumnMovesProps): PlannedColumnMove[] {
    const deleted = new Set(
      this._sentColumnTraits(sent, "deleted").map(({ columnId }) => columnId),
    );
    // One order per Table, so each move is indexed in the layout the earlier ones leave.
    const orders = new Map<string, TableColumnOrder>();
    const columnGroupMoves = groupMoves.map(({ table, columns, before }) => ({
      tableId: this.tableConfigOperator.tableIdOf(table),
      tableKey: table,
      columns,
      before,
    }));
    return columnGroupMoves.flatMap((ordered) => {
      const { tableId, tableKey } = ordered;
      const order =
        orders.get(tableId) ??
        this._columnOrderAfterChanges({ tableId, tableKey, deleted, minted });
      orders.set(tableId, order);
      return this._plannedMoves(order, ordered);
    });
  }
  // Last column first, each before the one after it in the group, so the group ends in its order before its target.
  private _plannedMoves(
    order: TableColumnOrder,
    { tableId, tableKey, columns, before }: OrderedColumns,
  ): PlannedColumnMove[] {
    const sequence = [...columns, before];
    return columns
      .map((column, index) => ({
        column,
        before: Val.assert(sequence[index + 1], `column after ${column}`),
      }))
      .reverse()
      .flatMap(({ column, before: target }): PlannedColumnMove[] => {
        const span = order.moveSpan({
          columnId: this._columnIdOf(tableKey, column),
          beforeColumnId: this._columnIdOf(tableKey, target),
          column,
          before: target,
        });
        if (span === undefined) return [];
        order.move(span);
        return [{ tableId, table: tableKey, ...span }];
      });
  }
  // Queued once the run's deletions and inserts have flushed, since their indexes are the layout those leave.
  moveColumns(moves: PlannedColumnMove[]): this {
    for (const { tableId, colIndex, count, beforeColIndex } of moves) {
      this.ss.raw
        .table(tableId)
        .moveColumns({ colIndex, count, beforeColIndex });
    }
    return this;
  }
  // As updateDeletionsAndInserts and appendColumns leave the Table: each insert before its column in entry order, each append at the end.
  private _columnOrderAfterChanges({
    tableId,
    tableKey,
    deleted,
    minted,
  }: ColumnOrderProps): TableColumnOrder {
    const order = this._columnOrder(tableId, tableKey).without(deleted);
    for (const {
      tableId: mintedTableId,
      table: key,
      columnId,
      before,
    } of minted) {
      if (mintedTableId !== tableId) continue;
      let beforeColumnId: string | undefined;
      if (before !== undefined) beforeColumnId = this._columnIdOf(key, before);
      order.insert(columnId, beforeColumnId);
    }
    return order;
  }
  private _columnOrder(tableId: string, tableKey: string): TableColumnOrder {
    const table = this.ss.raw.table(tableId);
    return new TableColumnOrder(
      tableKey,
      table.fullTableColIndexes.map((colIndex) =>
        table.columnResolver.columnIdAt(colIndex),
      ),
    );
  }
  private _colIndexOf(tableId: string, columnId: string): number {
    return this.ss.raw.table(tableId).columnResolver.colIndexOf(columnId);
  }
  private _isFirstColumn({
    tableId,
    columnId,
  }: Pick<NamedColumnIdentity, "tableId" | "columnId">): boolean {
    return this._colIndexOf(tableId, columnId) === 0;
  }
  private _columnIdOf(tableKey: string, columnName: string): string {
    return Val.assert(
      this._liveColumnNamed(tableKey, columnName),
      `live column ${tableKey}.${columnName}`,
    ).columnId;
  }
  // By previous names, as a `before` gives them.
  private _liveColumnNamed(
    tableKey: string,
    columnName: string,
  ): NamedColumnIdentity | undefined {
    return this.apiColumnIdentities
      .namedColumnIdentities()
      .find(
        (identity) =>
          identity.tableKey === tableKey && identity.columnName === columnName,
      );
  }
  private _sentColumnTraits(
    sent: IntendedTrait[],
    trait: "header" | "columnType" | "formula" | "deleted",
  ): SentColumnTrait[] {
    const identities = this.apiColumnIdentities.namedColumnIdentities();
    return sent.flatMap((intended) => {
      if (intended.kind !== "column" || intended.trait !== trait) return [];
      const { table: tableKey, column, value } = intended;
      const identity = Val.assert(
        identities.find(
          (named) => named.tableKey === tableKey && named.columnName === column,
        ),
        `live column ${tableKey}.${column}`,
      );
      return [{ ...identity, value }];
    });
  }
  recordEmptyValueAllowed(configTraits: TraitReadings): this {
    const byIds = this.columnConfigSync.emptyValueAllowedByIds;
    byIds.clear();
    for (const {
      tableId,
      columnId,
      tableKey,
      columnName,
    } of this.apiColumnIdentities.namedColumnIdentities()) {
      const value =
        configTraits.tables[tableKey]?.columns[columnName]?.emptyValueAllowed;
      if (typeof value !== "boolean") continue;
      const byColumnId = byIds.get(tableId) ?? new Map<string, boolean>();
      byColumnId.set(columnId, value);
      byIds.set(tableId, byColumnId);
    }
    return this;
  }
  untypedColumnsSummary(): string | undefined {
    return this.untypedColumnsOperator.untypedColumnsSummary();
  }
  newColumnConfigs(): ColumnConfigsGeneric {
    return this.outputOperator.newColumnConfigs();
  }
  apiTableColumnsByTableId(): Map<string, ApiTableColumns> {
    return this.outputOperator.apiTableColumnsByTableId();
  }
  identifiedApiTableColumnsByTableId(): Map<string, ApiTableColumns> {
    return this.outputOperator.identifiedApiTableColumnsByTableId();
  }
  toFileSource(): string {
    return this.outputOperator.toFileSource();
  }
}

// The names on the sheet when a new Table lands: its flush comes after the one sending Table renames and headers, and each planned Table answers to its own name and each column to the name its header gives.
export function namesAtTableCreation(
  sent: IntendedTrait[],
  newTables: PlannedNewTable[],
): BraceNames {
  const sentNames = namesAfterSends(sent);
  return newTables.reduce<BraceNames>(
    (withPlanned, { table, name, columns }) => {
      withPlanned.tableConfigs[table] = { tableName: name };
      withPlanned.columnConfigs[table] = Object.fromEntries(
        columns.map(({ column, header }) => [column, { header }]),
      );
      return withPlanned;
    },
    {
      tableConfigs: { ...sentNames.tableConfigs },
      columnConfigs: { ...sentNames.columnConfigs },
    },
  );
}

function namesAfterSends(sent: IntendedTrait[]): BraceNames {
  const { tableConfigs, columnConfigs } =
    new InstalledBraceConfigs().atColumnFill(sent);
  const reheaded = sent.reduce<BraceNames["columnConfigs"]>(
    (byKey, intended) => {
      if (intended.kind !== "column" || intended.trait !== "header") {
        return byKey;
      }
      const tableColumns = byKey[intended.table];
      if (tableColumns?.[intended.column] === undefined) return byKey;
      byKey[intended.table] = {
        ...tableColumns,
        [intended.column]: { header: String(intended.value) },
      };
      return byKey;
    },
    { ...columnConfigs },
  );
  return { tableConfigs, columnConfigs: reheaded };
}
