import type { SheetChange } from "../00_Source/PlatformEvents/sheetChange";
import type { ColumnName } from "../01_SpreadsheetSchema/columnConfigsTypes";
import {
  configSheetFloorSeed,
  type FloorSeedColumn,
  floorSeedColumns,
  type FloorTabName,
  floorTabSeedByGid,
} from "../01_SpreadsheetSchema/configSheetFloorSeed";
import { getSheetTraitByName } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import type { SheetRaw } from "../02_SpreadsheetRaw/SheetRaw";
import { SpreadsheetBaseNamed } from "../04_SpreadsheetNamed/ClassBases/SpreadsheetBaseNamed";
import type { ColumnNamed } from "../04_SpreadsheetNamed/ColumnNamed";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import { Obj } from "../utils/Obj";
import { Val } from "../utils/Val";
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
    this._assertFloorTitlesAreOwned();
    const presentFloorSheets = floorTabNames().flatMap((sheetName) => {
      const sheetGid = getSheetTraitByName(sheetName, "sheetGid");
      if (!this.ss.raw.gidIsActive(sheetGid)) return [];
      return [
        {
          sheet: this.ss.raw.sheet(sheetGid),
          seed: configSheetFloorSeed[sheetName],
        },
      ];
    });
    presentFloorSheets.forEach(({ sheet, seed }) => {
      assertFloorTable(sheet, seed.tableName);
    });
    const titleLines: string[] = [];
    const tableNameLines: string[] = [];
    presentFloorSheets.forEach(({ sheet, seed }) => {
      if (sheet.title !== seed.title) {
        titleLines.push(`"${sheet.title}" → ${seed.title}`);
        sheet.updateTitle(seed.title);
      }
      if (sheet.tables.length !== 1) return;
      const table = Val.assert(sheet.tables[0], "floor table");
      if (table.name === seed.tableName) return;
      tableNameLines.push(
        `${seed.title}'s Table "${table.name}" → ${seed.tableName}`,
      );
      sheet.updateTableName(seed.tableName);
    });
    return [
      ...reportLines("Restored tab titles", titleLines),
      ...reportLines("Restored Table names", tableNameLines),
    ];
  }
  private _assertFloorTitlesAreOwned(): void {
    const ownedGidByTitle = new Map<string, number>(
      floorTabNames().map((sheetName) => [
        configSheetFloorSeed[sheetName].title,
        getSheetTraitByName(sheetName, "sheetGid"),
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
    floorSheetNames().forEach((sheetName) => {
      const sheetGid = getSheetTraitByName(sheetName, "sheetGid");
      if (!this.ss.raw.gidIsActive(sheetGid)) return;
      const sheet = this.ss.sheet(sheetName);
      sheet.meta.uniformRow("columnId").prepFetchFull();
      sheet.meta.uniformRow("tableHeader").prepFetchFull();
      sheet.meta.uniformRow("colGroupName").prepFetchFull();
      if (floorDataValueColumns(sheetName).length > 0) {
        sheet.row(sheet.schema.topDataRowIdx).prepFetchFull();
      }
      sheet.prepFetchEditProtections();
    });
    this.ss.fetchAllPrepped({ includeProgrammaticFacts: true });
    return identityColIndexes;
  }
  private _ensureColumnLabels(): string[] {
    const headerLines: string[] = [];
    const columnIdLines: string[] = [];
    const groupHeadingLines: string[] = [];
    floorSheetNames().forEach((sheetName) => {
      const sheetGid = getSheetTraitByName(sheetName, "sheetGid");
      if (!this.ss.raw.gidIsActive(sheetGid)) return;
      const sheet = this.ss.sheet(sheetName);
      if (sheet.raw.tables.length !== 1) return;
      const meta = sheet.raw.meta;
      floorColumnsToRestore(sheetName).forEach((floorColumn) => {
        const colIndex = liveColIndex(meta, floorColumn);
        if (colIndex === undefined) return;
        const liveHeader = String(meta.tableHeaderRow.valueOrEmpty(colIndex));
        if (liveHeader !== floorColumn.header) {
          meta.tableHeaderRow.updateValue(colIndex, floorColumn.header);
          headerLines.push(
            `${sheet.raw.title} · ${liveHeader} (${floorColumn.columnId}) → ${floorColumn.header}`,
          );
        }
        const liveColumnId = String(meta.colIdRow.valueOrEmpty(colIndex));
        if (liveColumnId !== floorColumn.columnId) {
          meta.colIdRow.updateValue(colIndex, floorColumn.columnId);
          columnIdLines.push(
            `${sheet.raw.title} · ${floorColumn.header} (${liveColumnId}) → ${floorColumn.columnId}`,
          );
        }
        const liveHeading = String(
          meta.uniformRow("colGroupName").valueOrEmpty(colIndex),
        );
        if (liveHeading !== floorColumn.groupHeading) {
          meta
            .uniformRow("colGroupName")
            .updateValue(colIndex, floorColumn.groupHeading);
          const headingLabel =
            floorColumn.groupHeading === ""
              ? "(blank)"
              : floorColumn.groupHeading;
          groupHeadingLines.push(
            `${sheet.raw.title} · ${floorColumn.header} (${floorColumn.columnId}) → ${headingLabel}`,
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
    return floorSheetNames().flatMap((sheetName) =>
      this._ensureSheetDataValues(sheetName),
    );
  }
  private _ensureSheetDataValues<SN extends FloorSheetName>(
    sheetName: SN,
  ): string[] {
    const sheetGid = getSheetTraitByName(sheetName, "sheetGid");
    if (!this.ss.raw.gidIsActive(sheetGid)) return [];
    const sheet = this.ss.sheet(sheetName);
    if (sheet.raw.tables.length !== 1) return [];
    const row = sheet.raw.row(sheet.schema.topDataRowIdx);
    const restoredLines: string[] = [];
    floorDataValueColumns(sheetName).forEach((seedColumn) => {
      const colIndex = liveColIndex(
        sheet.raw.meta,
        floorColumnRestore(sheetName, {
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
    const typeChangeLines = floorSheetNames().flatMap((sheetName) =>
      this._ensureSheetColumnTypes(sheetName, floorSeedColumns(sheetName)),
    );
    return reportLines("Set column types", typeChangeLines);
  }
  private _ensureSheetColumnTypes<SN extends FloorSheetName>(
    sheetName: SN,
    columns: readonly FloorSeedColumn[],
  ): string[] {
    const sheetGid = getSheetTraitByName(sheetName, "sheetGid");
    if (!this.ss.raw.gidIsActive(sheetGid)) return [];
    const sheet = this.ss.sheet(sheetName);
    return columns.flatMap((seedColumn) => {
      const colIndex = liveColIndex(
        sheet.raw.meta,
        floorColumnRestore(sheetName, {
          header: seedColumn.header,
          groupHeading: "",
        }),
      );
      if (colIndex === undefined) return [];
      const column = sheet.column(
        columnNameByHeader(sheetName, seedColumn.header),
      );
      if (column.meta.activeColumnType === seedColumn.columnType) {
        return [];
      }
      column.meta.updateColumnType(seedColumn.columnType);
      return [`${floorColumnIdentity(column)} → ${seedColumn.columnType}`];
    });
  }
}

function floorTabNames(): FloorTabName[] {
  return Obj.keys(configSheetFloorSeed);
}

type FloorDataValueColumn = FloorSeedColumn & { dataValue: string };

function floorDataValueColumns(
  sheetName: FloorSheetName,
): FloorDataValueColumn[] {
  return floorSeedColumns(sheetName).filter(
    (seedColumn): seedColumn is FloorDataValueColumn =>
      seedColumn.dataValue !== undefined,
  );
}

function assertFloorTable(sheet: SheetRaw, tableName: string): void {
  if (sheet.tables.length === 0) {
    throw new Error(`Floor tab "${sheet.title}" has no Table.`);
  }
  if (
    sheet.tables.length > 1 &&
    !sheet.tables.some((table) => table.name === tableName)
  ) {
    throw new Error(
      `Floor tab "${sheet.title}" has several Tables and none is named ${tableName}.`,
    );
  }
}

function reportLines(label: string, lines: string[]): string[] {
  return lines.length > 0 ? [`${label}: ${lines.join("; ")}`] : [];
}

function floorColumnIdentity<
  SN extends FloorSheetName,
  CN extends ColumnName<SN>,
>(column: ColumnNamed<SN, CN>): string {
  const header = String(column.meta.uniformCell("tableHeader").valueOrEmpty());
  return `${column.sheet.raw.title} · ${header} (${column.columnId})`;
}
