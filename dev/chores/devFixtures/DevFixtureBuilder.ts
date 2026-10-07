import type { SheetColIndex } from "../../../src/00_Source/RawSource/SheetIndex";
import { dimensionIds } from "../../../src/01_SpreadsheetSchema/dimensionIds";
import { headRows } from "../../../src/01_SpreadsheetSchema/headRows";
import { getTableTraitByName } from "../../../src/01_SpreadsheetSchema/tableConfigsTypes";
import { TableOrigin } from "../../../src/01_SpreadsheetSchema/TableOrigin";
import { SpreadsheetBaseNamed } from "../../../src/04_SpreadsheetNamed/ClassBases/SpreadsheetBaseNamed";
import { SpreadsheetNamed } from "../../../src/04_SpreadsheetNamed/SpreadsheetNamed";
import {
  type DevFixtureColumn,
  type DevFixtureSheet,
  devFixtureSheets,
} from "./devFixtureSheets";

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
      this._ensureLetApiAccess(fixture);
      this._ensureEmptyValueAllowed(fixture);
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
    const { sheetGid, columns } = fixture;
    const origin = TableOrigin.expected();
    const rowCount = Math.max(...columns.map((column) => column.values.length));
    const endRowIdx = origin.sheetRowIndex(rowCount);
    const endColIdx = origin.sheetColIndex(columns.length);
    this.ss.raw
      .gatherAddSheetOperation({
        sheetId: sheetGid,
        title: fixture.title,
        rowCount: endRowIdx,
        columnCount: endColIdx,
      })
      .gatherAddTableOperation({
        // The live fixture Tables carry their name as their ID, so a rebuilt one matches its generated entry.
        tableId: fixture.tableName,
        name: fixture.tableName,
        range: {
          sheetId: sheetGid,
          startRowIndex: origin.headerRowIndex,
          endRowIndex: endRowIdx,
          startColumnIndex: origin.startColIndex,
          endColumnIndex: endColIdx,
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
        value: dimensionIds.col(fixture.idPrefix, column.key),
      });
      this._seedColumnRows(sheetGid, colIndex, column, rowCount);
    });
    if (fixture.entryCheckboxColumnKey !== undefined) {
      this._addEntryCheckbox(fixture, fixture.entryCheckboxColumnKey);
    }
  }
  private _seedColumnRows(
    sheetId: number,
    colIndex: SheetColIndex,
    column: DevFixtureColumn,
    rowCount: number,
  ): void {
    const origin = TableOrigin.expected();
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
  private _addEntryCheckbox(fixture: DevFixtureSheet, columnKey: string): void {
    const columnIndex = fixture.columns.findIndex(
      (column) => column.key === columnKey,
    );
    if (columnIndex === -1) {
      throw new Error(`${fixture.title} has no "${columnKey}" column.`);
    }
    const origin = TableOrigin.expected();
    const actionRowIndex = headRows.index("action");
    const rowIndex = origin.sheetRowIndex(actionRowIndex);
    const colIndex = origin.sheetColIndex(columnIndex);
    this.ss.raw
      .gatherAddedSheetFillCellOperation({
        sheetId: fixture.sheetGid,
        rowIndex,
        colIndex,
        value: false,
      })
      .gatherAddedSheetCheckboxValidationOperation({
        sheetId: fixture.sheetGid,
        startRowIndex: rowIndex,
        endRowIndex: origin.sheetRowIndex(actionRowIndex + 1),
        startColumnIndex: colIndex,
        endColumnIndex: origin.sheetColIndex(columnIndex + 1),
      });
  }
  private _ensureLetApiAccess(fixture: DevFixtureSheet): void {
    const tableConfig = this.ss.table("tableConfig");
    const [row] = tableConfig.rowsFiltered({ tableId: fixture.tableName });
    if (row === undefined) {
      tableConfig.appendRowWithVals({
        tableId: fixture.tableName,
        tableName: fixture.tableName,
        sheetTitle: fixture.title,
        letApiAccess: true,
      });
    } else if (row.valueOrEmpty("letApiAccess") !== true) {
      row.cell("letApiAccess").updateValue(true);
    }
  }
  private _ensureEmptyValueAllowed(fixture: DevFixtureSheet): void {
    const columnConfig = this.ss.table("columnConfig");
    fixture.columns.forEach(({ key, header, emptyValueAllowed }) => {
      if (emptyValueAllowed === undefined) return;
      const columnId = dimensionIds.col(fixture.idPrefix, key);
      const [row] = columnConfig.rowsFiltered({
        tableId: fixture.tableName,
        columnId,
      });
      if (row === undefined) {
        columnConfig.appendRowWithVals({
          tableId: fixture.tableName,
          columnId,
          tableName: fixture.tableName,
          header,
          emptyValueAllowed,
        });
      } else if (row.valueOrEmpty("emptyValueAllowed") !== emptyValueAllowed) {
        row.cell("emptyValueAllowed").updateValue(emptyValueAllowed);
      }
    });
  }
}
