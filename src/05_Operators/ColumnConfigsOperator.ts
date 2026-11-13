import { Str } from "@byronbroughten/utils/str";
import { Val } from "@byronbroughten/utils/val";

import { isFrameworkValueName } from "../00_Source/CellValues/frameworkValueSchemas";
import {
  type BraceConfigs,
  BraceFormula,
} from "../01_SpreadsheetSchema/configReaders/BraceFormula";
import { columnConfigsByName } from "../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import {
  type ColumnConfigsGeneric,
  type ColumnConfigStored,
  makeImportLine,
} from "../01_SpreadsheetSchema/configReaders/makeConfigs";
import {
  type IntendedTrait,
  type PlanReadings,
  type TableReading,
  type TraitReadings,
} from "../01_SpreadsheetSchema/configReaders/SyncPlanner";
import { tableConfigsByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { type ValueName } from "../01_SpreadsheetSchema/configReaders/valueSchemas";
import { dimensionIds } from "../01_SpreadsheetSchema/dimensionIds";
import type {
  ColumnProfileRaw,
  TableColumnReference,
} from "../02_SpreadsheetRaw/ColumnProfileRaw";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import { oneLinePerLeafFileSource } from "./configFileSource";
import { PreviousColumnConfigs } from "./PreviousColumnConfigs";
import {
  type ConfigSyncState,
  type SourceColumn,
  sourceLabel,
  SpreadsheetBaseOperator,
  type UntypedHeadersByTableName,
  type ValueSource,
} from "./SpreadsheetBaseOperator";
import { TableConfigsOperator } from "./TableConfigsOperator";

export interface ApiTableColumns {
  tableKey: string;
  columnConfigs: ColumnConfigsGeneric[string];
}

interface ColumnIdentity {
  tableId: string;
  columnId: string;
}

interface NamedColumnIdentity extends ColumnIdentity {
  tableKey: string;
  columnName: string;
  previousEmptyValueAllowed: boolean;
}

interface SentColumnTrait extends NamedColumnIdentity {
  value: IntendedTrait["value"];
}

export class ColumnConfigsOperator extends SpreadsheetBaseOperator {
  static init(): ColumnConfigsOperator {
    return new ColumnConfigsOperator(
      SpreadsheetBaseOperator.initOperatorProps(),
    );
  }
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  get columnConfigSync(): ConfigSyncState["columnConfigSync"] {
    return this.configSyncState.columnConfigSync;
  }
  get tableConfigOperator(): TableConfigsOperator {
    return new TableConfigsOperator(this.operatorProps);
  }
  private get untypedHeadersByTableName(): UntypedHeadersByTableName {
    return this.columnConfigSync.untypedHeadersByTableName;
  }
  // Derived fresh each call, not cached — a stored field goes stale across
  // this coordinator's per-access getter rebuilds.
  private get tableIdsApiAccesses(): Set<string> {
    return new Set(this.tableConfigOperator.tableIdsApiAccesses());
  }
  valueSources(): ValueSource[] {
    const tableKeysByTableId = this.tableConfigOperator.tableKeysByTableId();
    const sourcesByValueName = new Map<string, ValueSource>();
    for (const identity of this._columnIdentities()) {
      const source = this._valueSource(
        this._describedColumn(identity),
        tableKeysByTableId,
      );
      if (source === undefined) continue;
      const claimed = sourcesByValueName.get(source.valueName);
      if (claimed !== undefined && !isSameSourceColumn(claimed, source)) {
        throw new Error(
          `Value sources ${this._sourceLabel(claimed)} and ${this._sourceLabel(source)} ` +
            `both take the value name "${source.valueName}". Rename a Table or a header.`,
        );
      }
      sourcesByValueName.set(source.valueName, source);
    }
    return [...sourcesByValueName.values()];
  }
  private _sourceLabel({ tableId, header }: SourceColumn): string {
    return sourceLabel({ tableName: this.ss.raw.table(tableId).name, header });
  }
  private _valueSource(
    describedColumn: ColumnProfileRaw,
    tableKeysByTableId: Map<string, string>,
  ): ValueSource | undefined {
    const sourceColumn = this._sourceColumn(describedColumn);
    if (sourceColumn === undefined) return undefined;
    const { tableId, header } = sourceColumn;
    return {
      ...sourceColumn,
      valueName: valueSourceName({
        tableKey: Val.assert(
          tableKeysByTableId.get(tableId),
          "source Table key",
        ),
        reference: { tableName: this.ss.raw.table(tableId).name, header },
      }),
    };
  }
  private _sourceColumn(
    describedColumn: ColumnProfileRaw,
  ): SourceColumn | undefined {
    const reference = describedColumn.validationColumnReference();
    if (reference === undefined || describedColumn.isIdColumn) return undefined;
    const tableId = [...this.tableIdsApiAccesses].find(
      (apiTableId) =>
        this.ss.raw.table(apiTableId).name === reference.tableName,
    );
    if (tableId === undefined) return undefined;
    if (!this.ss.raw.table(tableId).profile.hasHeader(reference.header)) {
      return undefined;
    }
    return { tableId, header: reference.header };
  }
  assertSyncedToSpreadsheet(): void {
    if (!this.columnConfigSync.syncedToSpreadsheet) {
      throw new Error(
        "ColumnConfigsOperator has not yet synced to the spreadsheet.",
      );
    }
  }
  fetchHeadRows(): this {
    return this.fetchHeadRowsOf([...this.tableIdsApiAccesses]);
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
    for (const tableId of this.tableIdsApiAccesses) {
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
    this._recordUntypedColumns();
    this._logUntypedColumns();
    this.columnConfigSync.syncedToSpreadsheet = true;
    return this;
  }
  // emptyValueAllowed isn't a sheet trait, so its live side carries the previous configs and only a deleted column can stop it.
  traitReadings(): PlanReadings {
    const configs = installedBraceConfigs();
    const live = this._namedColumnIdentities().reduce<TraitReadings["tables"]>(
      (tables, identity) => {
        const { tableId, tableKey, columnName } = identity;
        const reading = tables[tableKey] ?? {
          traits: { name: this.ss.raw.table(tableId).name },
          columns: {},
        };
        const { header, topFormula } = this._describedColumn(identity);
        reading.columns[columnName] = {
          header,
          emptyValueAllowed: identity.previousEmptyValueAllowed,
          // In brace form against the previous configs, the names an alteration's formula uses.
          formula: liveFormula(configs, tableKey, topFormula),
        };
        tables[tableKey] = reading;
        return tables;
      },
      {},
    );
    return {
      previous: { tables: previousTraitTables(), tabs: {} },
      live: { tables: live, tabs: {} },
    };
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
    const configs = installedBraceConfigs();
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
  private _sentColumnTraits(
    sent: IntendedTrait[],
    trait: "header" | "formula",
  ): SentColumnTrait[] {
    const identities = this._namedColumnIdentities();
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
    } of this._namedColumnIdentities()) {
      const value =
        configTraits.tables[tableKey]?.columns[columnName]?.emptyValueAllowed;
      if (typeof value !== "boolean") continue;
      const byColumnId = byIds.get(tableId) ?? new Map<string, boolean>();
      byColumnId.set(columnId, value);
      byIds.set(tableId, byColumnId);
    }
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
  // Keyed by previous names through the IDs, so a rename keeps an alteration's target (ADR 0035).
  private _namedColumnIdentities(): NamedColumnIdentity[] {
    const tableKeysByTableId = this.tableConfigOperator.tableKeysByTableId();
    return this._columnIdentities().map((identity) => {
      const previous = new PreviousColumnConfigs(identity.tableId);
      return {
        ...identity,
        tableKey: Val.assert(
          previous.tableKey ?? tableKeysByTableId.get(identity.tableId),
          "API-access Table key",
        ),
        columnName:
          previous.columnName(identity.columnId) ??
          Str.sentenceToCamelCase(this._describedColumn(identity).header),
        previousEmptyValueAllowed: previous.emptyValueAllowed(
          identity.columnId,
        ),
      };
    });
  }
  // Per Table, the previous configs' order first, so a quiet run leaves the generated files unchanged.
  private _columnIdentities(): ColumnIdentity[] {
    return [...this.tableIdsApiAccesses].flatMap((tableId) => {
      const { profile } = this.ss.raw.table(tableId);
      const liveColumnIds = new Set(profile.columnIds);
      const previousColumnIds = new PreviousColumnConfigs(
        tableId,
      ).columnIds.filter((columnId) => liveColumnIds.has(columnId));
      const previous = new Set(previousColumnIds);
      const newColumnIds = profile.columnIds.filter(
        (columnId) => !previous.has(columnId),
      );
      return [...previousColumnIds, ...newColumnIds].map((columnId) => ({
        tableId,
        columnId,
      }));
    });
  }
  private _describedColumn({
    tableId,
    columnId,
  }: ColumnIdentity): ColumnProfileRaw {
    return this.ss.raw.table(tableId).profile.columnById(columnId);
  }
  private _recordUntypedColumns(): void {
    this.untypedHeadersByTableName.clear();
    for (const identity of this._columnIdentities()) {
      const describedColumn = this._describedColumn(identity);
      if (
        this._sourceColumn(describedColumn) === undefined &&
        describedColumn.declaredValueTitle() === undefined
      ) {
        this._recordUntypedColumn(
          this.ss.raw.table(identity.tableId).name,
          describedColumn.header,
        );
      }
    }
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
    const columnConfigs: ColumnConfigsGeneric = {};
    for (const identity of this._columnIdentities()) {
      const { tableId, columnId } = identity;
      const describedColumn = this._describedColumn(identity);
      const { header } = describedColumn;
      const tableKey = Val.assert(
        tableKeysByTableId.get(tableId),
        "API-access Table key",
      );
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
      const valueName =
        this._valueSource(describedColumn, tableKeysByTableId)?.valueName ??
        this.schema.titleToName(describedColumn.valueTitle());
      tableColumnConfigs[columnName] = {
        columnId,
        header,
        valueName: valueName as ValueName,
        isFormula: describedColumn.isFormula,
        emptyValueAllowed:
          this.columnConfigSync.emptyValueAllowedByIds
            .get(tableId)
            ?.get(columnId) ??
          new PreviousColumnConfigs(tableId).emptyValueAllowed(columnId),
        customDefaultValue: null,
        formula: describedColumn.topFormula ?? null,
        columnType: {
          type: describedColumn.columnType ?? null,
          rule: describedColumn.validationRule ?? null,
        },
        actionRowCell: describedColumn.actionRowCell ?? null,
        groupHeading1: describedColumn.groupHeading1 || null,
      };
    }
    return this._withFormulasInBraceForm(columnConfigs);
  }
  apiTableColumnsByTableId(): Map<string, ApiTableColumns> {
    const columnConfigs = this.newColumnConfigs();
    const tableKeysByTableId = this.tableConfigOperator.tableKeysByTableId();
    return [...tableKeysByTableId].reduce(
      (byTableId, [tableId, tableKey]) =>
        byTableId.set(tableId, {
          tableKey,
          columnConfigs: Val.assert(
            columnConfigs[tableKey],
            `column configs of ${tableKey}`,
          ),
        }),
      new Map<string, ApiTableColumns>(),
    );
  }
  // A second pass, because a brace names a column by the name this run gives it.
  private _withFormulasInBraceForm(
    columnConfigs: ColumnConfigsGeneric,
  ): ColumnConfigsGeneric {
    const configs = {
      tableConfigs: this.tableConfigOperator.newTableConfigs(),
      columnConfigs,
    };
    return Object.entries(columnConfigs).reduce<ColumnConfigsGeneric>(
      (braced, [tableKey, tableColumnConfigs]) => {
        braced[tableKey] = Object.entries(tableColumnConfigs).reduce<
          ColumnConfigsGeneric[string]
        >((bracedTable, [columnName, columnConfig]) => {
          bracedTable[columnName] = {
            ...columnConfig,
            formula: bracedFormula({
              configs,
              tableKey,
              formula: columnConfig.formula,
            }),
          };
          return bracedTable;
        }, {});
        return braced;
      },
      {},
    );
  }
  toFileSource(): string {
    return [
      `${makeImportLine(["makeColumnConfigs"])}`,
      ``,
      `export const columnConfigs = makeColumnConfigs(${oneLinePerLeafFileSource(
        this.newColumnConfigs(),
      )});`,
      ``,
    ].join("\n");
  }
}

function previousTraitTables(): TraitReadings["tables"] {
  const tableConfigs = tableConfigsByName();
  return Object.entries(columnConfigsByName()).reduce<TraitReadings["tables"]>(
    (tables, [tableKey, columns]) => {
      const tableName = tableConfigs[tableKey]?.tableName;
      tables[tableKey] = {
        traits: tableName === undefined ? {} : { name: tableName },
        columns: Object.entries(columns).reduce<TableReading["columns"]>(
          (byName, [columnName, { header, emptyValueAllowed, formula }]) => {
            byName[columnName] = {
              header,
              emptyValueAllowed,
              formula: formula ?? undefined,
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

function installedBraceConfigs(): BraceConfigs {
  return {
    tableConfigs: tableConfigsByName(),
    columnConfigs: columnConfigsByName(),
  };
}

function bracedFormula({
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

function liveFormula(
  configs: BraceConfigs,
  tableKey: string,
  topFormula: string | undefined,
): string | undefined {
  if (topFormula === undefined) return undefined;
  return BraceFormula.decompile({ configs, tableKey, formula: topFormula });
}

function valueSourceName({
  tableKey,
  reference,
}: {
  tableKey: string;
  reference: TableColumnReference;
}): string {
  const headerName = Str.sentenceToCamelCase(reference.header);
  const valueName =
    tableKey + headerName.charAt(0).toUpperCase() + headerName.slice(1);
  if (isFrameworkValueName(valueName)) {
    throw new Error(
      `Value source ${sourceLabel(reference)} takes the value name "${valueName}", ` +
        `which is already the framework value name "${valueName}". Rename the Table or the header.`,
    );
  }
  return valueName;
}

function isSameSourceColumn(a: SourceColumn, b: SourceColumn): boolean {
  return a.tableId === b.tableId && a.header === b.header;
}
