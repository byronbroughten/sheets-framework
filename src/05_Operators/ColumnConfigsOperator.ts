import {
  type BraceConfigs,
  BraceFormula,
  type BraceNames,
} from "../01_SpreadsheetSchema/configReaders/BraceFormula.ts";
import { columnNameOfHeader } from "../01_SpreadsheetSchema/configReaders/columnConfigsTypes.ts";
import type { ColumnConfigsGeneric } from "../01_SpreadsheetSchema/configReaders/makeConfigs.ts";
import type { FormulaColumn } from "../01_SpreadsheetSchema/configReaders/OutsideReferences.ts";
import { previousTraitTables } from "../01_SpreadsheetSchema/configReaders/previousTraitTables.ts";
import {
  type IntendedTrait,
  newColumnLines,
  newTableLines,
  type PlannedNewColumn,
  type PlannedNewTable,
  type PlanReadings,
  type TraitReadings,
} from "../01_SpreadsheetSchema/configReaders/SyncPlanner.ts";
import type {
  SyncBeforeFirstColumnStop,
  SyncCreatedColumn,
  SyncCreatedTable,
  SyncHeaderTakenStop,
  SyncSendLine,
} from "../01_SpreadsheetSchema/configReaders/syncRunResult.ts";
import { dimensionIds } from "../01_SpreadsheetSchema/dimensionIds.ts";
import { ColumnConfigsBaseOperator } from "./ClassBases/ColumnConfigsBaseOperator.ts";
import type { NamedColumnIdentity } from "./ColumnConfigsOperator/ApiColumnIdentities.ts";
import {
  ColumnConfigsLayoutOperator,
  type ColumnMovesProps,
  type MintedNewColumn,
  type PlannedColumnMove,
} from "./ColumnConfigsOperator/ColumnConfigsLayoutOperator.ts";
import {
  type ApiTableColumns,
  ColumnConfigsOutputOperator,
} from "./ColumnConfigsOperator/ColumnConfigsOutputOperator.ts";
import { ColumnConfigsUntypedColumnsOperator } from "./ColumnConfigsOperator/ColumnConfigsUntypedColumnsOperator.ts";
import { ColumnConfigsValueSourcesOperator } from "./ColumnConfigsOperator/ColumnConfigsValueSourcesOperator.ts";
import {
  bracedFormula,
  InstalledBraceConfigs,
} from "./ColumnConfigsOperator/InstalledBraceConfigs.ts";
import {
  SpreadsheetBaseOperator,
  type ValueSource,
} from "./SpreadsheetBaseOperator.ts";
import type { MintedNewTable } from "./TableConfigsOperator.ts";

export interface LandedColumns {
  createdColumns: SyncCreatedColumn[];
  didntLand: SyncSendLine[];
}

export interface LandedTables {
  createdTables: SyncCreatedTable[];
  didntLand: SyncSendLine[];
}

export class ColumnConfigsOperator extends ColumnConfigsBaseOperator {
  static init(): ColumnConfigsOperator {
    return new ColumnConfigsOperator(
      SpreadsheetBaseOperator.initOperatorProps(),
    );
  }
  get valueSourcesOperator(): ColumnConfigsValueSourcesOperator {
    return new ColumnConfigsValueSourcesOperator(this.operatorProps);
  }
  get untypedColumnsOperator(): ColumnConfigsUntypedColumnsOperator {
    return new ColumnConfigsUntypedColumnsOperator(this.operatorProps);
  }
  get outputOperator(): ColumnConfigsOutputOperator {
    return new ColumnConfigsOutputOperator(this.operatorProps);
  }
  get layoutOperator(): ColumnConfigsLayoutOperator {
    return new ColumnConfigsLayoutOperator(this.operatorProps);
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
        const next = this.layoutOperator
          ._columnOrder(tableId, tableKey)
          .columnIdAfter(columnId);
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
          .columnBodyA1(this.apiColumnIdentities.colIndexOf(tableId, columnId));
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
    for (const {
      tableId,
      columnId,
      value,
    } of this.apiColumnIdentities.sentColumnTraits(sent, "header")) {
      this.ss.raw.table(tableId).updateHeader(columnId, String(value));
    }
    return this;
  }
  // Fills every body row with one formula; a Table copies a uniform column formula into the rows it grows by.
  updateFormulas(sent: IntendedTrait[]): this {
    const configs = this.installedBraceConfigs.atColumnFill(sent);
    for (const {
      tableId,
      columnId,
      tableKey,
      value,
    } of this.apiColumnIdentities.sentColumnTraits(sent, "formula")) {
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
    for (const {
      tableId,
      columnId,
      tableKey,
      value,
    } of this.apiColumnIdentities.sentColumnTraits(sent, "columnType")) {
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
    const { apiColumnIdentities } = this;
    return newColumns.flatMap(
      ({ table, beforePath, header, before }): SyncBeforeFirstColumnStop[] => {
        if (before === undefined) return [];
        const placedBefore = apiColumnIdentities.identityNamed(table, before);
        if (placedBefore === undefined) return [];
        if (!apiColumnIdentities.isFirstColumn(placedBefore)) return [];
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
  mintColumnIds(newColumns: PlannedNewColumn[]): MintedNewColumn[] {
    return this.layoutOperator.mintColumnIds(newColumns);
  }
  appendColumns(minted: MintedNewColumn[]): this {
    this.layoutOperator.appendColumns(minted);
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
  updateDeletionsAndInserts(
    sent: IntendedTrait[],
    minted: MintedNewColumn[],
  ): this {
    this.layoutOperator.updateDeletionsAndInserts(sent, minted);
    return this;
  }
  planColumnMoves(props: ColumnMovesProps): PlannedColumnMove[] {
    return this.layoutOperator.planColumnMoves(props);
  }
  moveColumns(moves: PlannedColumnMove[]): this {
    this.layoutOperator.moveColumns(moves);
    return this;
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
