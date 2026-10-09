import { Obj } from "@byronbroughten/utils/obj";

import type { SheetChange } from "../00_Source/PlatformEvents/sheetChange";
import type { ColumnName } from "../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import {
  configSheetFloorSeed,
  type FloorSeedColumn,
  floorSeedColumns,
  type FloorTabName,
  floorTabSeedByGid,
} from "../01_SpreadsheetSchema/configReaders/configSheetFloorSeed";
import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import type { TableRaw } from "../02_SpreadsheetRaw/TableRaw";
import { SpreadsheetBaseNamed } from "../04_SpreadsheetNamed/ClassBases/SpreadsheetBaseNamed";
import type { ColumnNamed } from "../04_SpreadsheetNamed/ColumnNamed";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import { ConfigSheetFloorCreator } from "./ConfigSheetFloor/ConfigSheetFloorCreator";
import {
  ConfigSheetFloorEditWarnings,
  type IdentityColIndexes,
} from "./ConfigSheetFloor/ConfigSheetFloorEditWarnings";
import {
  floorChangeNotice,
  type FloorNotice,
} from "./ConfigSheetFloor/floorChangeNotice";
import { liveColIndex } from "./ConfigSheetFloor/floorColumnLocation";
import {
  columnNameByHeader,
  floorColumnRestore,
  floorColumnsToRestore,
  type FloorSheetName,
  floorSheetNames,
} from "./ConfigSheetFloor/floorSeedLookups";

export const retiredSheetConfigTitle = "Sheet Config";

/**
 * Restores floor tab titles, Table names, headers, column IDs, group
 * headings, data values and column types, has ConfigSheetFloorCreator
 * create missing floor tabs and columns, and has ConfigSheetFloorEditWarnings declare
 * the edit warnings. ConfigCoordinator
 * runs this at the start of every config sync; the
 * ensureConfigSheetFloor chore is the other caller.
 * changeNotice has floorChangeNotice decide the floor notice for an On change
 * event that renamed or deleted Value Config, the one floor tab with no edit warning.
 * docs/generated-data/config-sheet-floor.md
 */
export class ConfigSheetFloor extends SpreadsheetBaseNamed {
  static init(): ConfigSheetFloor {
    return new ConfigSheetFloor(ConfigSheetFloor.initSpreadsheetNamedProps());
  }
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  private get editWarnings(): ConfigSheetFloorEditWarnings {
    return new ConfigSheetFloorEditWarnings(this.spreadsheetNamedProps);
  }
  private get creator(): ConfigSheetFloorCreator {
    return new ConfigSheetFloorCreator(this.spreadsheetNamedProps);
  }
  ensure(): string {
    const titleAndTableLines = this._ensureTitlesAndTables();
    let identityColIndexes = this._fetchFloorSheets();
    const createdLines = this.creator.createMissing();
    if (createdLines.length > 0) {
      this.ss.batchUpdateGSheets();
      this.ss.fetchAllSheetProperties();
      identityColIndexes = this._fetchFloorSheets();
    }
    return [
      ...titleAndTableLines,
      ...createdLines,
      ...this._ensureColumnLabels(),
      ...this._ensureDataValues(),
      ...this._ensureColumnTypes(),
      ...this.editWarnings.ensure(identityColIndexes),
    ].join("; ");
  }
  changeNotice(change: SheetChange): FloorNotice | undefined {
    this.ss.raw.fetchAllSheetProperties();
    const liveTitlesByGid = new Map(
      this.ss.raw.activeSheetGids.flatMap((sheetGid) =>
        floorTabSeedByGid(sheetGid) === undefined
          ? []
          : [[sheetGid, this.ss.raw.sheet(sheetGid).title] as const],
      ),
    );
    return floorChangeNotice(change, liveTitlesByGid);
  }
  private _ensureTitlesAndTables(): string[] {
    this.ss.raw.ensureAllSheetPropertiesAreFetched();
    this._assertSheetConfigIsConverted();
    this._assertFloorTitlesAreOwned();
    const presentFloorSheets = floorTabNames().flatMap((tableName) => {
      const sheetGid = getTableTraitByName(tableName, "sheetGid");
      if (!this.ss.raw.gidIsActive(sheetGid)) return [];
      const floorTableId = getTableTraitByName(tableName, "tableId");
      return [
        {
          sheet: this.ss.raw.table(floorTableId, sheetGid),
          seed: configSheetFloorSeed[tableName],
          floorTableId,
        },
      ];
    });
    presentFloorSheets.forEach(({ sheet, floorTableId }) => {
      assertOnlyFloorTable(sheet, floorTableId);
    });
    const titleLines: string[] = [];
    const tableNameLines: string[] = [];
    presentFloorSheets.forEach(({ sheet, seed }) => {
      if (sheet.sheet.title !== seed.title) {
        titleLines.push(`"${sheet.sheet.title}" → ${seed.title}`);
        sheet.sheet.updateTitle(seed.title);
      }
      if (sheet.name === seed.liveTableName) return;
      tableNameLines.push(
        `${seed.title}'s Table "${sheet.name}" → ${seed.liveTableName}`,
      );
      sheet.updateTableName(seed.liveTableName);
    });
    return [
      ...reportLines("Restored tab titles", titleLines),
      ...reportLines("Restored Table names", tableNameLines),
    ];
  }
  // Restoring an unconverted Sheet Config tab as Table Config would drop its ticks.
  private _assertSheetConfigIsConverted(): void {
    const titles = this.ss.raw.activeSheetGids.map(
      (sheetGid) => this.ss.raw.sheet(sheetGid).title,
    );
    if (
      titles.includes(retiredSheetConfigTitle) &&
      !titles.includes(configSheetFloorSeed.tableConfig.title)
    ) {
      throw new Error(
        `Found a "${retiredSheetConfigTitle}" tab and no "${configSheetFloorSeed.tableConfig.title}" tab. Run the convertSheetConfigToTableConfig chore before syncing or regenerating configs.`,
      );
    }
  }
  private _assertFloorTitlesAreOwned(): void {
    const ownedGidByTitle = new Map<string, number>(
      floorTabNames().map((tableName) => [
        configSheetFloorSeed[tableName].title,
        getTableTraitByName(tableName, "sheetGid"),
      ]),
    );
    this.ss.raw.activeSheetGids.forEach((sheetGid) => {
      const title = this.ss.raw.sheet(sheetGid).title;
      const ownedGid = ownedGidByTitle.get(title);
      if (ownedGid !== undefined && sheetGid !== ownedGid) {
        throw new Error(`A tab titled "${title}" is not the floor tab.`);
      }
    });
  }
  private _fetchFloorSheets(): IdentityColIndexes {
    const identityColIndexes = this.editWarnings.gatherIdentityColumns();
    floorSheetNames().forEach((tableName) => {
      const sheetGid = getTableTraitByName(tableName, "sheetGid");
      if (!this.ss.raw.gidIsActive(sheetGid)) return;
      const sheet = this.ss.table(tableName);
      sheet.headRow("columnId").prepFetchFull();
      sheet.headRow("header").prepFetchFull();
      sheet.headRow("groupHeading1").prepFetchFull();
      if (floorDataValueColumns(tableName).length > 0) {
        sheet.row(0).prepFetchFull();
      }
      sheet.sheet.prepFetchEditProtections();
    });
    this.ss.fetchAllPrepped({ includeProgrammaticFacts: true });
    return identityColIndexes;
  }
  private _ensureColumnLabels(): string[] {
    const headerLines: string[] = [];
    const columnIdLines: string[] = [];
    const groupHeadingLines: string[] = [];
    floorSheetNames().forEach((tableName) => {
      const sheetGid = getTableTraitByName(tableName, "sheetGid");
      if (!this.ss.raw.gidIsActive(sheetGid)) return;
      const sheet = this.ss.table(tableName);
      const rawTable = sheet.raw;
      const headerRow = rawTable.headRow("header");
      const colIdRow = rawTable.headRow("columnId");
      floorColumnsToRestore(tableName).forEach((floorColumn) => {
        const colIndex = liveColIndex(rawTable, floorColumn);
        if (colIndex === undefined) return;
        const liveHeader = String(headerRow.valueOrEmpty(colIndex));
        if (liveHeader !== floorColumn.header) {
          headerRow.updateValue(colIndex, floorColumn.header);
          headerLines.push(
            `${sheet.raw.sheet.title} · ${liveHeader} (${floorColumn.columnId}) → ${floorColumn.header}`,
          );
        }
        const liveColumnId = String(colIdRow.valueOrEmpty(colIndex));
        if (liveColumnId !== floorColumn.columnId) {
          colIdRow.updateValue(colIndex, floorColumn.columnId);
          columnIdLines.push(
            `${sheet.raw.sheet.title} · ${floorColumn.header} (${liveColumnId}) → ${floorColumn.columnId}`,
          );
        }
        const liveHeading = String(
          rawTable.headRow("groupHeading1").valueOrEmpty(colIndex),
        );
        if (liveHeading !== floorColumn.groupHeading) {
          rawTable
            .headRow("groupHeading1")
            .updateValue(colIndex, floorColumn.groupHeading);
          const headingLabel =
            floorColumn.groupHeading === ""
              ? "(blank)"
              : floorColumn.groupHeading;
          groupHeadingLines.push(
            `${sheet.raw.sheet.title} · ${floorColumn.header} (${floorColumn.columnId}) → ${headingLabel}`,
          );
        }
      });
    });
    return [
      ...reportLines("Restored headers", headerLines),
      ...reportLines("Restored column IDs", columnIdLines),
      ...reportLines("Restored group headings", groupHeadingLines),
    ];
  }
  private _ensureDataValues(): string[] {
    return floorSheetNames().flatMap((tableName) =>
      this._ensureSheetDataValues(tableName),
    );
  }
  private _ensureSheetDataValues<TN extends FloorSheetName>(
    tableName: TN,
  ): string[] {
    const sheetGid = getTableTraitByName(tableName, "sheetGid");
    if (!this.ss.raw.gidIsActive(sheetGid)) return [];
    const sheet = this.ss.table(tableName);
    const row = sheet.raw.row(0);
    const restoredLines: string[] = [];
    floorDataValueColumns(tableName).forEach((seedColumn) => {
      const colIndex = liveColIndex(
        sheet.raw,
        floorColumnRestore(tableName, {
          header: seedColumn.header,
          groupHeading: "",
        }),
      );
      if (colIndex === undefined) return;
      const liveValue = String(row.cell(colIndex).valueOrEmpty());
      const { dataValue } = seedColumn;
      if (liveValue === dataValue) return;
      row.updateValue(colIndex, dataValue);
      restoredLines.push(
        `Restored ${seedColumn.header}: "${liveValue}" → ${dataValue}`,
      );
    });
    return restoredLines;
  }
  private _ensureColumnTypes(): string[] {
    const typeChangeLines = floorSheetNames().flatMap((tableName) =>
      this._ensureSheetColumnTypes(tableName, floorSeedColumns(tableName)),
    );
    return reportLines("Set column types", typeChangeLines);
  }
  private _ensureSheetColumnTypes<TN extends FloorSheetName>(
    tableName: TN,
    columns: readonly FloorSeedColumn[],
  ): string[] {
    const sheetGid = getTableTraitByName(tableName, "sheetGid");
    if (!this.ss.raw.gidIsActive(sheetGid)) return [];
    const sheet = this.ss.table(tableName);
    return columns.flatMap((seedColumn) => {
      const colIndex = liveColIndex(
        sheet.raw,
        floorColumnRestore(tableName, {
          header: seedColumn.header,
          groupHeading: "",
        }),
      );
      if (colIndex === undefined) return [];
      const column = sheet.column(
        columnNameByHeader(tableName, seedColumn.header),
      );
      if (column.raw.profile.columnType === seedColumn.columnType) {
        return [];
      }
      column.updateColumnType(seedColumn.columnType);
      return [`${floorColumnIdentity(column)} → ${seedColumn.columnType}`];
    });
  }
}

function floorTabNames(): FloorTabName[] {
  return Obj.keys(configSheetFloorSeed);
}

type FloorDataValueColumn = FloorSeedColumn & { dataValue: string };

function floorDataValueColumns(
  tableName: FloorSheetName,
): FloorDataValueColumn[] {
  return floorSeedColumns(tableName).filter(
    (seedColumn): seedColumn is FloorDataValueColumn =>
      seedColumn.dataValue !== undefined,
  );
}

function assertOnlyFloorTable(sheet: TableRaw, floorTableId: string): void {
  const tableIds = sheet.tableIds();
  if (tableIds.length === 0) {
    throw new Error(`${floorTabLabel(sheet)} has no Table.`);
  }
  if (tableIds.length === 1 && tableIds[0] === floorTableId) return;
  if (tableIds.length === 1) {
    throw new Error(
      `${floorTabLabel(sheet)} has ${tableNamesLabel(sheet, tableIds)}, not its floor Table (Table ID "${floorTableId}"); regenerate the configs if it replaced the floor Table.`,
    );
  }
  if (!tableIds.includes(floorTableId)) {
    throw new Error(
      `${floorTabLabel(sheet)} has ${tableNamesLabel(sheet, tableIds)} and none is its floor Table (Table ID "${floorTableId}").`,
    );
  }
  const extraTableIds = tableIds.filter((tableId) => tableId !== floorTableId);
  throw new Error(
    `${floorTabLabel(sheet)} holds only its floor Table; move or delete ${tableNamesLabel(sheet, extraTableIds)}.`,
  );
}

function floorTabLabel(sheet: TableRaw): string {
  return `Floor tab "${sheet.sheet.title}"`;
}

function tableNamesLabel(sheet: TableRaw, tableIds: string[]): string {
  return tableIds
    .map((tableId) => `Table "${sheet.ss.table(tableId).name}"`)
    .join(", ");
}

function reportLines(label: string, lines: string[]): string[] {
  return lines.length > 0 ? [`${label}: ${lines.join("; ")}`] : [];
}

function floorColumnIdentity<
  TN extends FloorSheetName,
  CN extends ColumnName<TN>,
>(column: ColumnNamed<TN, CN>): string {
  const header = String(column.headCell("header").valueOrEmpty());
  return `${column.table.raw.sheet.title} · ${header} (${column.columnId})`;
}
