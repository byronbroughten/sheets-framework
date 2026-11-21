import { Val } from "@byronbroughten/utils/val";

import type { ModelableConditionalFormatRule } from "../00_Source/RawSource/ConditionalFormat";
import type { BoundedGridRange } from "../00_Source/RawSource/RawSource";
import {
  type ApiAccessChanges,
  type ApiAccessGrant,
  hasApiAccess,
} from "../01_SpreadsheetSchema/configReaders/apiAccessChanges";
import {
  BraceFormula,
  type BraceNames,
} from "../01_SpreadsheetSchema/configReaders/BraceFormula";
import {
  type ColumnReference,
  type IdNameMapGeneric,
  mapTablesByLiveness,
} from "../01_SpreadsheetSchema/configReaders/idNameMapTypes";
import { IntendedRule } from "../01_SpreadsheetSchema/configReaders/IntendedRule";
import {
  makeImportLine,
  type TableConfigsBase,
} from "../01_SpreadsheetSchema/configReaders/makeConfigs";
import type {
  IntendedTrait,
  PlannedNewTable,
  PlannedNewTableColumn,
  PlannedNewTablePlacement,
} from "../01_SpreadsheetSchema/configReaders/SyncPlanner";
import type {
  SyncHeadRowsStop,
  SyncNewTableStop,
  SyncTableGoneNotice,
  SyncTabTitleRepeatedStop,
} from "../01_SpreadsheetSchema/configReaders/syncRunResult";
import {
  configRetiredIdPrefixes,
  previousTableIdOf,
  tableConfigsByTableId,
} from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { dimensionIds } from "../01_SpreadsheetSchema/dimensionIds";
import { headerZone } from "../01_SpreadsheetSchema/headerZone";
import {
  idPrefixes,
  IdPrefixesByTableId,
  type IdPrefixLabel,
} from "../01_SpreadsheetSchema/idPrefixes";
import { TableOrigin } from "../01_SpreadsheetSchema/TableOrigin";
import type { CreatedTablePlacement } from "../02_SpreadsheetRaw/ClassTypes/StateRaw";
import type { SheetRaw } from "../02_SpreadsheetRaw/SheetRaw";
import {
  randomSheetGid,
  randomTableId,
} from "../02_SpreadsheetRaw/SpreadsheetRaw/SpreadsheetTableCreatorRaw";
import type { TableRaw } from "../02_SpreadsheetRaw/TableRaw";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import { oneLinePerEntryFileSource } from "./configFileSource";
import { validateIdPrefixesUnchanged } from "./idPrefixGuard";
import { SpreadsheetBaseOperator } from "./SpreadsheetBaseOperator";
import { TableConfigsTraitUpdaterOperator } from "./TableConfigsOperator/TableConfigsTraitUpdaterOperator";

// A planned new Table with the IDs the run settled for it, its tab's GID in its created placement, and each column's formula and type compiled to the names on the sheet.
export interface MintedNewTable extends PlannedNewTable {
  tableId: string;
  createdPlacement: CreatedTablePlacement;
  idPrefix: string;
  columns: (PlannedNewTableColumn & { columnId: string })[];
}

export class TableConfigsOperator extends SpreadsheetBaseOperator {
  static init(): TableConfigsOperator {
    return new TableConfigsOperator(
      SpreadsheetBaseOperator.initOperatorProps(),
    );
  }
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  get traitUpdater(): TableConfigsTraitUpdaterOperator {
    return new TableConfigsTraitUpdaterOperator(this.operatorProps);
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
  // A grant writes head rows above the header, so its header must sit where they fit, clear of every other Table.
  grantHeadRowStops(grants: ApiAccessGrant[]): SyncHeadRowsStop[] {
    const { raw } = this.ss;
    return grants.flatMap(({ tableId, table }): SyncHeadRowsStop[] => {
      if (!raw.tableIdIsActive(tableId)) return [];
      const granted = raw.table(tableId);
      const { startRowIndex } = granted;
      const tab = raw.sheet(granted.sheetGid).title;
      if (!headerZone.holdsHeaderRow(startRowIndex)) {
        return [{ kind: "headRows", table, tab, headerRow: startRowIndex + 1 }];
      }
      return granted.neighboursUnderHeadRows.map((neighbour) => ({
        kind: "headRows",
        table,
        tab,
        neighbour: neighbour.name,
      }));
    });
  }
  // Matched by the codebase name the regen would give, since two Tables or tabs giving one name can't both be configured.
  newTableStops(newTables: PlannedNewTable[]): SyncNewTableStop[] {
    const { raw } = this.ss;
    const liveTables = raw.activeTableIds.map((tableId) => raw.table(tableId));
    const perTable = newTables.flatMap(({ table, name, placement }) => {
      const stops: SyncNewTableStop[] = [];
      const takenTable = liveTables.find(
        (live) => this.schema.titleToName(live.name) === table,
      );
      if (takenTable !== undefined) {
        stops.push({
          kind: "tableNameTaken",
          name,
          existing: takenTable.name,
          tableId: takenTable.tableId,
        });
      }
      if (placement.kind !== "newTab") return stops;
      const { title } = placement;
      const takenTab = this._liveTabByName(this.schema.titleToName(title));
      if (takenTab !== undefined) {
        stops.push({
          kind: "tabTitleTaken",
          title,
          existing: takenTab.title,
          gid: takenTab.sheetGid,
        });
      }
      return stops;
    });
    return [...perTable, ...this._tabTitleRepeatedStops(newTables)];
  }
  // Since #331 a new tab's Tables are placed by its title, so a repeated title would quietly merge two entries into one tab, and a case-only difference would make two tabs with one codebase name.
  private _tabTitleRepeatedStops(
    newTables: PlannedNewTable[],
  ): SyncTabTitleRepeatedStop[] {
    const pathsByName = new Map<string, { title: string; paths: string[] }>();
    for (const { placement } of newTables) {
      if (placement.kind !== "newTab") continue;
      const name = this.schema.titleToName(placement.title);
      const entry = pathsByName.get(name) ?? {
        title: placement.title,
        paths: [],
      };
      if (!entry.paths.includes(placement.path)) {
        entry.paths.push(placement.path);
      }
      pathsByName.set(name, entry);
    }
    return [...pathsByName.values()].flatMap(
      ({ title, paths }): SyncTabTitleRepeatedStop[] => {
        if (paths.length < 2) return [];
        return [{ kind: "tabTitleRepeated", title, paths }];
      },
    );
  }
  // Minted before any queue, so a dry run's withheld flush and the report carry the same IDs; compiled then too, so a bad brace throws before any flush.
  mintNewTables(
    newTables: PlannedNewTable[],
    names: BraceNames,
  ): MintedNewTable[] {
    const prefixesInUse = new Set(this.idPrefixesByTableId().idPrefixes());
    const retiredPrefixes = this._retiredOrGoneIdPrefixes();
    const sheetGidsInUse = new Set(this.ss.raw.activeSheetGids);
    const newTabGidsByTitle = new Map<string, number>();
    return newTables.map((newTable) => {
      const idPrefix = idPrefixes.fromTitle({
        title: newTable.name,
        prefixesInUse,
        retiredPrefixes,
      });
      prefixesInUse.add(idPrefix);
      const createdPlacement = this._createdPlacement({
        placement: newTable.placement,
        sheetGidsInUse,
        newTabGidsByTitle,
      });
      sheetGidsInUse.add(createdPlacement.sheetGid);
      if (createdPlacement.kind === "newTab") {
        newTabGidsByTitle.set(
          createdPlacement.title,
          createdPlacement.sheetGid,
        );
      }
      const mintedTable = {
        ...newTable,
        tableId: randomTableId(),
        createdPlacement,
        idPrefix,
        columns: newTable.columns.map((column) => ({
          ...BraceFormula.compileColumn({
            configs: names,
            tableKey: newTable.table,
            column,
          }),
          columnId: dimensionIds.col(idPrefix),
        })),
      };
      // Thrown away: only a rule's position waits on where the Table lands, and its colour, condition and braces don't.
      compiledRules(mintedTable, provisionalRange(mintedTable));
      return mintedTable;
    });
  }
  private _createdPlacement({
    placement,
    sheetGidsInUse,
    newTabGidsByTitle,
  }: CreatedPlacementProps): CreatedTablePlacement {
    if (placement.kind === "newTab") {
      const newTabGid = newTabGidsByTitle.get(placement.title);
      if (newTabGid !== undefined) {
        return { kind: "besideTables", sheetGid: newTabGid };
      }
      return {
        kind: "newTab",
        sheetGid: randomSheetGid(sheetGidsInUse),
        title: placement.title,
      };
    }
    const tab = Val.assert(
      this._liveTabByName(placement.tab),
      `live tab ${placement.tab}`,
    );
    return { kind: "besideTables", sheetGid: tab.sheetGid, gap: placement.gap };
  }
  private _liveTabByName(tab: string): SheetRaw | undefined {
    const { raw } = this.ss;
    return raw.activeSheetGids
      .map((sheetGid) => raw.sheet(sheetGid))
      .find((live) => this.schema.titleToName(live.title) === tab);
  }
  // Each through the shared creation path; the run flushes them after every column write, which a beside placement can't follow.
  createTables(minted: MintedNewTable[]): this {
    for (const newTable of minted) {
      const { range } = this.ss.raw.createTable({
        name: newTable.name,
        placement: newTable.createdPlacement,
        idPrefix: newTable.idPrefix,
        tableId: newTable.tableId,
        columns: newTable.columns.map(
          ({ header, columnId, columnType, formula }) => ({
            header,
            columnId,
            columnType,
            formula,
          }),
        ),
      });
      // Reversed, as each lands first, so the list's first rule ends on top.
      for (const compiled of compiledRules(newTable, range).reverse()) {
        this.ss.raw.gatherAddedTableConditionalFormatRule(compiled);
      }
    }
    return this;
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
  goneTableNotices(previousIdNameMap: IdNameMapGeneric): SyncTableGoneNotice[] {
    const { gone } = mapTablesByLiveness(
      previousIdNameMap,
      new Set(this.ss.raw.activeTableIds),
    );
    return gone.filter(hasApiAccess).map(({ name }): SyncTableGoneNotice => ({
      kind: "tableGone",
      table: name,
    }));
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
  updateTableNames(sent: IntendedTrait[]): this {
    this.traitUpdater.updateTableNames(sent);
    return this;
  }
  deleteTables(sent: IntendedTrait[]): this {
    this.traitUpdater.deleteTables(sent);
    return this;
  }
  tableIdOf(tableKey: string): string {
    return previousTableIdOf(tableKey);
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

interface CreatedPlacementProps {
  placement: PlannedNewTablePlacement;
  sheetGidsInUse: Set<number>;
  newTabGidsByTitle: Map<string, number>;
}

// Each of a new Table's rules over the body rows of `range`, its columns left to right as its headers give them.
function compiledRules(
  { table, columns, rules }: MintedNewTable,
  range: BoundedGridRange,
): ModelableConditionalFormatRule[] {
  const layout = {
    headerTopLeft: {
      rowIndex: range.startRowIndex,
      colIndex: range.startColumnIndex,
    },
    endRowIndex: range.endRowIndex,
    columnNamesLeftToRight: columns.map(({ column }) => column),
  };
  return rules.map((rule) =>
    IntendedRule.compile({
      tableKey: table,
      sheetGid: range.sheetId,
      layout,
      rule,
    }),
  );
}

// Where the Table would sit alone on a tab, one body row deep.
function provisionalRange({
  createdPlacement,
  columns,
}: MintedNewTable): BoundedGridRange {
  const origin = TableOrigin.expected();
  return {
    sheetId: createdPlacement.sheetGid,
    startRowIndex: origin.headerRowIndex,
    endRowIndex: origin.sheetRowIndex(1),
    startColumnIndex: origin.startColIndex,
    endColumnIndex: origin.sheetColIndex(columns.length),
  };
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
