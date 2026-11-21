import type { CellValue } from "../../../src/00_Source/CellValues/cellValues";
import { dimensionIds } from "../../../src/01_SpreadsheetSchema/dimensionIds";
import { headRows } from "../../../src/01_SpreadsheetSchema/headRows";
import { type TableOrigin } from "../../../src/01_SpreadsheetSchema/TableOrigin";
import type {
  CreatedTablePlacement,
  CreateTableProps,
} from "../../../src/02_SpreadsheetRaw/ClassTypes/StateRaw";
import { SpreadsheetBaseNamed } from "../../../src/04_SpreadsheetNamed/ClassBases/SpreadsheetBaseNamed";
import { SpreadsheetNamed } from "../../../src/04_SpreadsheetNamed/SpreadsheetNamed";
import {
  bodyRowCountOf,
  type DevFixtureSheet,
  devFixtureSheets,
  type DevFixtureTable,
} from "./devFixtureSheets";

// Each tab's first Table opens it at the expected origin; the rest sit one blank column right of the Table before.
const besideGap = 1;

interface FirstTableExtrasProps {
  sheetGid: number;
  origin: TableOrigin;
  table: DevFixtureTable;
}

// A tab that exists is left alone; rebuild one by deleting it and rerunning (docs/how-it-runs/dev-project.md).
export class DevFixtureBuilder extends SpreadsheetBaseNamed {
  static init(): DevFixtureBuilder {
    return new DevFixtureBuilder(DevFixtureBuilder.initSpreadsheetNamedProps());
  }
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  ensureFixtures(): string {
    this.ss.fetchAllSheetProperties();
    const missing = devFixtureSheets.filter(
      (fixture) => !this.ss.raw.gidIsActive(fixture.sheetGid),
    );
    const unplanned = missing.flatMap((fixture) =>
      this._addFixtureSheet(fixture),
    );
    this.ss.batchUpdateGSheets();
    if (missing.length === 0) {
      return "Every fixture tab already exists; no Tables were planned.";
    }
    const created = `Created ${missing.map((fixture) => fixture.title).join(", ")}.`;
    if (unplanned.length === 0) return created;
    return `${created} Not planned, as their tab wasn't sent (a dry run): ${unplanned.join(", ")}.`;
  }
  // Returns the names of the Tables it couldn't place.
  private _addFixtureSheet(fixture: DevFixtureSheet): string[] {
    const { sheetGid, title, tables } = fixture;
    for (const [tableIndex, table] of tables.entries()) {
      if (tableIndex === 0) {
        const { origin } = this.ss.raw.createTable(
          createTablePropsOf(table, { kind: "newTab", sheetGid, title }),
        );
        this._addFirstTableExtras({ sheetGid, origin, table });
        continue;
      }
      validateNoFirstTableExtras(table);
      // Placing beside reads the tab's Tables, so the tab and the Tables before must be sent first.
      this.ss.batchUpdateGSheets();
      this.ss.fetchAllSheetProperties();
      if (!this.ss.raw.gidIsActive(sheetGid)) {
        return tables.slice(tableIndex).map(({ tableName }) => tableName);
      }
      this.ss.raw.createTable(
        createTablePropsOf(table, {
          kind: "besideTables",
          sheetGid,
          gap: besideGap,
        }),
      );
    }
    return [];
  }
  // Formulas and the entry checkbox need the Table's origin, which only a tab's first Table knows before the flush.
  private _addFirstTableExtras(props: FirstTableExtrasProps): void {
    const { sheetGid, origin, table } = props;
    for (const [columnIndex, { formula }] of table.columns.entries()) {
      if (formula === undefined) continue;
      for (let rowIndex = 0; rowIndex < bodyRowCountOf(table); rowIndex++) {
        this.ss.raw.gatherAddedTableFillCellOperation({
          sheetId: sheetGid,
          rowIndex: origin.sheetRowIndex(rowIndex),
          colIndex: origin.sheetColIndex(columnIndex),
          formula,
        });
      }
    }
    this._addEntryCheckbox(props);
  }
  private _addEntryCheckbox({
    sheetGid,
    origin,
    table,
  }: FirstTableExtrasProps): void {
    const columnKey = table.entryCheckboxColumnKey;
    if (columnKey === undefined) return;
    const columnIndex = table.columns.findIndex(
      (column) => column.key === columnKey,
    );
    if (columnIndex === -1) {
      throw new Error(`${table.tableName} has no "${columnKey}" column.`);
    }
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
}

// The live fixture Tables carry their name as their ID, so a rebuilt one matches its generated entry and keeps API access.
function createTablePropsOf(
  table: DevFixtureTable,
  placement: CreatedTablePlacement,
): CreateTableProps {
  const { tableName, idPrefix, columns } = table;
  return {
    name: tableName,
    placement,
    idPrefix,
    tableId: tableName,
    columns: columns.map(({ key, header, columnType }) => ({
      header,
      columnId: dimensionIds.col(idPrefix, key),
      columnType: { type: columnType, rule: null },
    })),
    seedRows: seedRowsOf(table),
  };
}

function seedRowsOf(table: DevFixtureTable): CellValue[][] {
  return Array.from({ length: bodyRowCountOf(table) }, (_, rowIndex) =>
    table.columns.map(({ values }) => values[rowIndex] ?? ""),
  );
}

function validateNoFirstTableExtras(table: DevFixtureTable): void {
  const hasFormula = table.columns.some(({ formula }) => formula !== undefined);
  if (hasFormula || table.entryCheckboxColumnKey !== undefined) {
    throw new Error(
      `${table.tableName} isn't its tab's first Table, so it can't take a formula or an entry checkbox.`,
    );
  }
}
