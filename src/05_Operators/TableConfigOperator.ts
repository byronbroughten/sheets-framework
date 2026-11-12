import { Val } from "@byronbroughten/utils/val";

import {
  makeImportLine,
  type TableConfigsBase,
} from "../01_SpreadsheetSchema/configReaders/makeConfigs";
import {
  configRetiredIdPrefixes,
  tableConfigsByTableId,
} from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import {
  idPrefixes,
  type IdPrefixLabel,
} from "../01_SpreadsheetSchema/idPrefixes";
import type { TableRaw } from "../02_SpreadsheetRaw/TableRaw";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import { oneLinePerEntryFileSource } from "./configFileSource";
import { validateIdPrefixesUnchanged } from "./idPrefixGuard";
import { SpreadsheetBaseOperator } from "./SpreadsheetBaseOperator";

export interface ColumnReference {
  tableId: string;
  columnId: string;
}

export class TableConfigOperator extends SpreadsheetBaseOperator {
  static init(): TableConfigOperator {
    return new TableConfigOperator(SpreadsheetBaseOperator.initOperatorProps());
  }
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  // The previous configs' order, so a quiet run leaves the generated files unchanged.
  tableIdsApiAccesses(): string[] {
    return [...tableConfigsByTableId().keys()].filter((tableId) =>
      this.ss.raw.tableIdIsActive(tableId),
    );
  }
  validateHeadRowsClear(): void {
    this.ss.raw.validateHeadRowsClear(this.tableIdsApiAccesses());
  }
  validateHeadersInZone(): void {
    this.ss.raw.validateHeadersInZone(this.tableIdsApiAccesses());
  }
  idPrefix(tableId: string): string {
    return Val.assert(this._idPrefixesByTableId().get(tableId), "ID prefix");
  }
  validateIdPrefixesUnchanged(): void {
    validateIdPrefixesUnchanged({
      previousByTableId: previousIdPrefixesByTableId(),
      nextByTableId: this._idPrefixesByTableId(),
      retiredPrefixes: this._retiredOrGoneIdPrefixes(),
    });
  }
  private _idPrefixesByTableId(): Map<string, string> {
    const previousByTableId = previousIdPrefixesByTableId();
    const assigned = new Map<string, string>();
    for (const table of this._apiAccessTables()) {
      const kept =
        table.profile.idPrefix() ?? previousByTableId.get(table.tableId);
      if (kept !== undefined) assigned.set(table.tableId, kept);
    }
    const prefixesInUse = new Set(assigned.values());
    const retiredPrefixes = this._retiredOrGoneIdPrefixes();
    for (const table of this._apiAccessTables()) {
      if (assigned.has(table.tableId)) continue;
      const generated = idPrefixes.fromTitle({
        title: table.name,
        prefixesInUse,
        retiredPrefixes,
      });
      prefixesInUse.add(generated);
      assigned.set(table.tableId, generated);
    }
    return assigned;
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
      currentByTableId: this._idPrefixesByTableId(),
    });
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
  tableKeysByTableId(): Map<string, string> {
    const map = new Map<string, string>();
    Object.entries(this.newTableConfigs()).forEach(([tableKey, config]) => {
      map.set(config.tableId, tableKey);
    });
    return map;
  }
  toFileSource(makeConfigsImport: string): string {
    return [
      makeImportLine(["makeTableConfigs"], makeConfigsImport),
      ``,
      `export const tableConfigs = makeTableConfigs(${oneLinePerEntryFileSource(
        this.newTableConfigs(),
      )});`,
      ``,
    ].join("\n");
  }
  retiredIdPrefixesFileSource(makeConfigsImport: string): string {
    return [
      makeImportLine(["makeRetiredIdPrefixes"], makeConfigsImport),
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

function previousIdPrefixesByTableId(): Map<string, string> {
  return new Map(
    [...tableConfigsByTableId()].map(([tableId, config]) => [
      tableId,
      config.idPrefix,
    ]),
  );
}
