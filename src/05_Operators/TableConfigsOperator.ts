import { Val } from "@byronbroughten/utils/val";

import {
  type ApiAccessChanges,
  type ApiAccessGrant,
  hasApiAccess,
} from "../01_SpreadsheetSchema/configReaders/apiAccessChanges";
import type {
  ColumnReference,
  IdNameMapGeneric,
} from "../01_SpreadsheetSchema/configReaders/idNameMapTypes";
import {
  makeImportLine,
  type TableConfigsBase,
} from "../01_SpreadsheetSchema/configReaders/makeConfigs";
import type { IntendedTrait } from "../01_SpreadsheetSchema/configReaders/SyncPlanner";
import type {
  SyncHeadRowsStop,
  SyncNotice,
} from "../01_SpreadsheetSchema/configReaders/syncRunResult";
import {
  configRetiredIdPrefixes,
  tableConfigsByName,
  tableConfigsByTableId,
} from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { headerZone } from "../01_SpreadsheetSchema/headerZone";
import {
  idPrefixes,
  IdPrefixesByTableId,
  type IdPrefixLabel,
} from "../01_SpreadsheetSchema/idPrefixes";
import type { TableRaw } from "../02_SpreadsheetRaw/TableRaw";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import { oneLinePerEntryFileSource } from "./configFileSource";
import { validateIdPrefixesUnchanged } from "./idPrefixGuard";
import { SpreadsheetBaseOperator } from "./SpreadsheetBaseOperator";

export class TableConfigsOperator extends SpreadsheetBaseOperator {
  static init(): TableConfigsOperator {
    return new TableConfigsOperator(
      SpreadsheetBaseOperator.initOperatorProps(),
    );
  }
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  // The previous configs' order, then this run's grants, so a quiet run leaves the generated files unchanged.
  tableIdsApiAccesses(): string[] {
    const { grants, revokedTableIds } = this.configSyncState.apiAccessSync;
    const revoked = new Set(revokedTableIds);
    return [
      ...tableConfigsByTableId().keys(),
      ...grants.map(({ tableId }) => tableId),
    ].filter(
      (tableId) =>
        !revoked.has(tableId) && this.ss.raw.tableIdIsActive(tableId),
    );
  }
  recordApiAccessChanges(changes: ApiAccessChanges): this {
    this.configSyncState.apiAccessSync = changes;
    return this;
  }
  // A grant writes head rows above the header, so its header must sit where they fit.
  grantHeadRowStops(grants: ApiAccessGrant[]): SyncHeadRowsStop[] {
    const { raw } = this.ss;
    return grants.flatMap(({ tableId, table }): SyncHeadRowsStop[] => {
      if (!raw.tableIdIsActive(tableId)) return [];
      const { startRowIndex, sheetGid } = raw.table(tableId);
      if (headerZone.holdsHeaderRow(startRowIndex)) return [];
      return [
        {
          kind: "headRows",
          table,
          tab: raw.sheet(sheetGid).title,
          headerRow: startRowIndex + 1,
        },
      ];
    });
  }
  validateHeadRowsClear(): void {
    this.ss.raw.validateHeadRowsClear(this.tableIdsApiAccesses());
  }
  validateHeadersInZone(): void {
    this.ss.raw.validateHeadersInZone(this.tableIdsApiAccesses());
  }
  validateIdPrefixesUnchanged(): void {
    validateIdPrefixesUnchanged({
      previousByTableId: previousIdPrefixesByTableId(),
      nextByTableId: this.idPrefixesByTableId(),
      retiredPrefixes: this._retiredOrGoneIdPrefixes(),
    });
  }
  idPrefixesByTableId(): IdPrefixesByTableId {
    const previousByTableId = previousIdPrefixesByTableId();
    const assigned = new Map<string, string>();
    const apiAccessTables = this._apiAccessTables();
    for (const table of apiAccessTables) {
      const kept =
        table.profile.idPrefix() ?? previousByTableId.get(table.tableId);
      if (kept !== undefined) assigned.set(table.tableId, kept);
    }
    const prefixesInUse = new Set(assigned.values());
    const retiredPrefixes = this._retiredOrGoneIdPrefixes();
    for (const table of apiAccessTables) {
      if (assigned.has(table.tableId)) continue;
      const generated = idPrefixes.fromTitle({
        title: table.name,
        prefixesInUse,
        retiredPrefixes,
      });
      prefixesInUse.add(generated);
      assigned.set(table.tableId, generated);
    }
    return new IdPrefixesByTableId(assigned);
  }
  // Read before any send, so a Table this run deletes is still live and only a person's deletion shows.
  goneTableNotices(previousIdNameMap: IdNameMapGeneric): SyncNotice[] {
    return Object.entries(previousIdNameMap.tables).flatMap(
      ([tableId, mapTable]): SyncNotice[] => {
        if (!hasApiAccess(mapTable)) return [];
        if (this.ss.raw.tableIdIsActive(tableId)) return [];
        return [{ kind: "tableGone", table: mapTable.name }];
      },
    );
  }
  // A Table gone since the last run retires its prefix before this run mints any.
  private _retiredOrGoneIdPrefixes(): Set<string> {
    return idPrefixes.retiredOrGone({
      previouslyRetired: configRetiredIdPrefixes(),
      previousByTableId: previousIdPrefixesByTableId(),
      liveTableIds: new Set(this.tableIdsApiAccesses()),
    });
  }
  retiredIdPrefixes(): string[] {
    return idPrefixes.retiredAfter({
      previouslyRetired: configRetiredIdPrefixes(),
      previousByTableId: previousIdPrefixesByTableId(),
      currentByTableId: this.idPrefixesByTableId(),
    });
  }
  newTableConfigs(): TableConfigsBase {
    const tableConfigs: TableConfigsBase = {};
    const idPrefixLabels: IdPrefixLabel[] = [];
    const idPrefixesByTableId = this.idPrefixesByTableId();
    this._apiAccessTables().forEach((table) => {
      const { tableId, name: tableName } = table;
      const tableKey = this.schema.titleToName(tableName);
      const existing = tableConfigs[tableKey];
      if (existing !== undefined) {
        throw new Error(
          `Tables "${existing.tableName}" and "${tableName}" both give the key "${tableKey}".`,
        );
      }
      const idPrefix = idPrefixesByTableId.assertIdPrefix(tableId);
      const headerRow = table.headRow("header");
      tableConfigs[tableKey] = {
        tableId,
        tableName,
        sheetGid: table.sheetGid,
        idPrefix,
        hasIdColumn: headerRow.hasValue(this.schema.idHeader),
        hasNameColumn: headerRow.hasValue(this.schema.nameHeader),
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
    const columnId = table.columnResolver.columnIdByHeader(header);
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
  // A renamed Table is addressed by its previous key, which the refuser guarantees the previous configs hold.
  updateTableNames(sent: IntendedTrait[]): this {
    for (const intended of sent) {
      if (intended.kind !== "table" || intended.trait !== "name") continue;
      this._previousTable(intended.table).updateTableName(
        String(intended.value),
      );
    }
    return this;
  }
  deleteTables(sent: IntendedTrait[]): this {
    for (const intended of sent) {
      if (intended.kind !== "table" || intended.trait !== "deleted") continue;
      this._previousTable(intended.table).queueDelete();
    }
    return this;
  }
  private _previousTable(tableKey: string): TableRaw {
    const { tableId } = Val.assert(
      tableConfigsByName()[tableKey],
      `previous Table ${tableKey}`,
    );
    return this.ss.raw.table(tableId);
  }
  tableKeysByTableId(): Map<string, string> {
    const map = new Map<string, string>();
    Object.entries(this.newTableConfigs()).forEach(([tableKey, config]) => {
      map.set(config.tableId, tableKey);
    });
    return map;
  }
  toFileSource(): string {
    return [
      makeImportLine(["makeTableConfigs"]),
      ``,
      `export const tableConfigs = makeTableConfigs(${oneLinePerEntryFileSource(
        this.newTableConfigs(),
      )});`,
      ``,
    ].join("\n");
  }
  retiredIdPrefixesFileSource(): string {
    return [
      makeImportLine(["makeRetiredIdPrefixes"]),
      ``,
      `export const retiredIdPrefixes = makeRetiredIdPrefixes(${JSON.stringify(
        this.retiredIdPrefixes(),
      )});`,
      ``,
    ].join("\n");
  }
}

function columnReferenceLabel(reference: string): string {
  return `Column reference "${reference}"`;
}

function previousIdPrefixesByTableId(): IdPrefixesByTableId {
  return new IdPrefixesByTableId(
    new Map(
      [...tableConfigsByTableId()].map(([tableId, config]) => [
        tableId,
        config.idPrefix,
      ]),
    ),
  );
}
