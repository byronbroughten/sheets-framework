import type {
  SheetColIndex,
  SheetRowIndex,
} from "../../../src/00_Source/RawSource/SheetIndex";
import { getTableTraitByName } from "../../../src/01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { dimensionIds } from "../../../src/01_SpreadsheetSchema/dimensionIds";
import { headRows } from "../../../src/01_SpreadsheetSchema/headRows";
import type { TableOrigin } from "../../../src/01_SpreadsheetSchema/TableOrigin";
import { SpreadsheetBaseNamed } from "../../../src/04_SpreadsheetNamed/ClassBases/SpreadsheetBaseNamed";
import { SpreadsheetNamed } from "../../../src/04_SpreadsheetNamed/SpreadsheetNamed";
import {
  bodyRowCountOf,
  type DevFixtureColumn,
  type DevFixtureSheet,
  devFixtureSheets,
  type DevFixtureTable,
} from "./devFixtureSheets";

interface FixtureColumnSeed {
  sheetId: number;
  origin: TableOrigin;
  colIndex: SheetColIndex;
  column: DevFixtureColumn;
  rowCount: number;
}

// A tab that exists is left alone; rebuild one by deleting it and rerunning (docs/how-it-runs.md).
export class DevFixtureBuilder extends SpreadsheetBaseNamed {
  static init(): DevFixtureBuilder {
    return new DevFixtureBuilder(DevFixtureBuilder.initSpreadsheetNamedProps());
  }
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  ensureFixtures(): string {
    this.ss.fetchAllSheetProperties();
    this._validateConfigFloorPresent();
    this.ss
      .table("tableConfig")
      .prepFetchColumnsFull("tableId", "letApiAccess");
    this.ss
      .table("columnConfig")
      .prepFetchColumnsFull("tableId", "columnId", "emptyValueAllowed");
    this.ss.fetchAllPrepped({ skipFetchingProperties: true });
    const missing = devFixtureSheets.filter(
      (fixture) => !this.ss.raw.gidIsActive(fixture.sheetGid),
    );
    missing.forEach((fixture) => this._addFixtureSheet(fixture));
    devFixtureSheets.forEach((fixture) => {
      fixture.tables.forEach((table) => {
        this._ensureLetApiAccess(fixture.title, table);
        this._ensureEmptyValueAllowed(table);
      });
    });
    this.ss.batchUpdateGSheets();
    if (missing.length === 0) {
      return "Every fixture tab already exists; only the config ticks were checked.";
    }
    return `Created ${missing.map((fixture) => fixture.title).join(", ")}.`;
  }
  private _validateConfigFloorPresent(): void {
    const missing = (["tableConfig", "columnConfig"] as const).filter(
      (sheetName) =>
        !this.ss.raw.gidIsActive(getTableTraitByName(sheetName, "sheetGid")),
    );
    if (missing.length > 0) {
      throw new Error(
        `No ${missing.join(" or ")} tab yet. Run \`npm run dev:gen:configs\` first to create the config floor.`,
      );
    }
  }
  private _addFixtureSheet(fixture: DevFixtureSheet): void {
    const { sheetGid, tables } = fixture;
    this.ss.raw.gatherAddSheetOperation({
      sheetId: sheetGid,
      title: fixture.title,
      rowCount: Math.max(...tables.map(endRowIndexOf)),
      columnCount: Math.max(...tables.map(endColIndexOf)),
    });
    tables.forEach((table) => this._addFixtureTable(sheetGid, table));
  }
  private _addFixtureTable(sheetGid: number, table: DevFixtureTable): void {
    const { origin, columns } = table;
    const rowCount = bodyRowCountOf(table);
    this.ss.raw.gatherAddTableOperation({
      // The live fixture Tables carry their name as their ID, so a rebuilt one matches its generated entry.
      tableId: table.tableName,
      name: table.tableName,
      range: {
        sheetId: sheetGid,
        startRowIndex: origin.headerRowIndex,
        endRowIndex: endRowIndexOf(table),
        startColumnIndex: origin.startColIndex,
        endColumnIndex: endColIndexOf(table),
      },
      columnProperties: columns.map((column, columnIndex) => ({
        columnIndex,
        columnName: column.header,
        columnType: column.columnType,
      })),
    });
    columns.forEach((column, columnIndex) => {
      const colIndex = origin.sheetColIndex(columnIndex);
      this.ss.raw.gatherAddedSheetFillCellOperation({
        sheetId: sheetGid,
        rowIndex: origin.headSheetRowIndex("columnId"),
        colIndex,
        value: dimensionIds.col(table.idPrefix, column.key),
      });
      this._seedColumnRows({
        sheetId: sheetGid,
        origin,
        colIndex,
        column,
        rowCount,
      });
    });
    this._addEntryCheckbox(sheetGid, table);
  }
  private _seedColumnRows({
    sheetId,
    origin,
    colIndex,
    column,
    rowCount,
  }: FixtureColumnSeed): void {
    for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
      const position = {
        sheetId,
        rowIndex: origin.sheetRowIndex(rowIndex),
        colIndex,
      };
      const { formula } = column;
      const value = column.values[rowIndex];
      if (formula !== undefined) {
        this.ss.raw.gatherAddedSheetFillCellOperation({ ...position, formula });
      } else if (value !== undefined && value !== "") {
        this.ss.raw.gatherAddedSheetFillCellOperation({ ...position, value });
      }
    }
  }
  private _addEntryCheckbox(sheetGid: number, table: DevFixtureTable): void {
    const columnKey = table.entryCheckboxColumnKey;
    if (columnKey === undefined) return;
    const columnIndex = table.columns.findIndex(
      (column) => column.key === columnKey,
    );
    if (columnIndex === -1) {
      throw new Error(`${table.tableName} has no "${columnKey}" column.`);
    }
    const { origin } = table;
    const actionRowIndex = headRows.index("action");
    const rowIndex = origin.sheetRowIndex(actionRowIndex);
    const colIndex = origin.sheetColIndex(columnIndex);
    this.ss.raw
      .gatherAddedSheetFillCellOperation({
        sheetId: sheetGid,
        rowIndex,
        colIndex,
        value: false,
      })
      .gatherAddedSheetCheckboxValidationOperation({
        sheetId: sheetGid,
        startRowIndex: rowIndex,
        endRowIndex: origin.sheetRowIndex(actionRowIndex + 1),
        startColumnIndex: colIndex,
        endColumnIndex: origin.sheetColIndex(columnIndex + 1),
      });
  }
  private _ensureLetApiAccess(
    sheetTitle: string,
    table: DevFixtureTable,
  ): void {
    const tableConfig = this.ss.table("tableConfig");
    const [row] = tableConfig.rowsFiltered({ tableId: table.tableName });
    if (row === undefined) {
      tableConfig.appendRowWithVals({
        tableId: table.tableName,
        tableName: table.tableName,
        sheetTitle,
        letApiAccess: true,
      });
    } else if (row.valueOrEmpty("letApiAccess") !== true) {
      row.cell("letApiAccess").updateValue(true);
    }
  }
  private _ensureEmptyValueAllowed(table: DevFixtureTable): void {
    const columnConfig = this.ss.table("columnConfig");
    table.columns.forEach(({ key, header, emptyValueAllowed }) => {
      if (emptyValueAllowed === undefined) return;
      const columnId = dimensionIds.col(table.idPrefix, key);
      const [row] = columnConfig.rowsFiltered({
        tableId: table.tableName,
        columnId,
      });
      if (row === undefined) {
        columnConfig.appendRowWithVals({
          tableId: table.tableName,
          columnId,
          tableName: table.tableName,
          header,
          emptyValueAllowed,
        });
      } else if (row.valueOrEmpty("emptyValueAllowed") !== emptyValueAllowed) {
        row.cell("emptyValueAllowed").updateValue(emptyValueAllowed);
      }
    });
  }
}

function endRowIndexOf(table: DevFixtureTable): SheetRowIndex {
  return table.origin.sheetRowIndex(bodyRowCountOf(table));
}

function endColIndexOf(table: DevFixtureTable): SheetColIndex {
  return table.origin.sheetColIndex(table.columns.length);
}
