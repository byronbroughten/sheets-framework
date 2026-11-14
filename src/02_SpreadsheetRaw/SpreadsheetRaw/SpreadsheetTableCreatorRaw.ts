import type { CellValue } from "../../00_Source/CellValues/cellValues";
import type { TableColumnPropertiesAdd } from "../../00_Source/RawSource/RawSource";
import {
  type SheetColIndex,
  SheetIndex,
  type SheetRowIndex,
} from "../../00_Source/RawSource/SheetIndex";
import type { ColumnTypeBaseline } from "../../01_SpreadsheetSchema/columnBaseline";
import { dimensionIds } from "../../01_SpreadsheetSchema/dimensionIds";
import { idPrefixes } from "../../01_SpreadsheetSchema/idPrefixes";
import { TableOrigin } from "../../01_SpreadsheetSchema/TableOrigin";
import { SpreadsheetBaseRaw } from "../ClassBases/SpreadsheetBaseRaw";
import type {
  CreatedTable,
  CreatedTablePlacement,
  CreateTableProps,
} from "../ClassTypes/StateRaw";
import { SpreadsheetRaw } from "../SpreadsheetRaw";

// A created Table holds one empty body row unless it is seeded.
const minBodyRowCount = 1;

export class SpreadsheetTableCreatorRaw extends SpreadsheetBaseRaw {
  get ss(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
  // IDs are settled before any queue, so a dry run's withheld flush and the caller's record carry the same ones.
  createTable(props: CreateTableProps): CreatedTable {
    const { name, placement, columns, seedRows = [] } = props;
    validateSeedRows(name, columns.length, seedRows);
    const created = createdIds(props);
    const origin = this._originFor(name, placement);
    const endRowIndex = origin.sheetRowIndex(
      Math.max(minBodyRowCount, seedRows.length),
    );
    const endColumnIndex = origin.sheetColIndex(columns.length);
    this._ensureGridFor(placement, { endRowIndex, endColumnIndex });
    this.ss.gatherAddTableOperation({
      tableId: created.tableId,
      name,
      range: {
        sheetId: placement.sheetGid,
        startRowIndex: origin.headerRowIndex,
        endRowIndex,
        startColumnIndex: origin.startColIndex,
        endColumnIndex,
      },
      columnProperties: columns.map(({ header, columnType }, columnIndex) => ({
        columnIndex,
        columnName: header,
        ...columnTypeProperties(columnType),
      })),
    });
    this._fillCells(placement.sheetGid, origin, [
      {
        rowIndex: origin.headSheetRowIndex("columnId"),
        values: created.columnIds,
      },
      ...seedRows.map((values, rowIndex) => ({
        rowIndex: origin.sheetRowIndex(rowIndex),
        values,
      })),
    ]);
    return created;
  }
  private _originFor(
    name: string,
    placement: CreatedTablePlacement,
  ): TableOrigin {
    const expected = TableOrigin.expected();
    if (placement.kind === "newTab") return expected;
    const rightmostEnd = this._rightmostTableEnd(name, placement.sheetGid);
    if (rightmostEnd === undefined) return expected;
    return new TableOrigin({
      headerRowIndex: expected.headerRowIndex,
      startColIndex: SheetIndex.col(rightmostEnd + (placement.gap ?? 0)),
    });
  }
  // A queued column write moves a Table's end at the flush, after this measures it.
  private _rightmostTableEnd(
    name: string,
    sheetGid: number,
  ): number | undefined {
    if (!this.ss.gidIsActive(sheetGid)) {
      throw new Error(
        `Table "${name}" can't be placed beside the Tables on GID ${sheetGid}: no fetched tab has it.`,
      );
    }
    const sheet = this.ss.sheet(sheetGid);
    sheet.assertRowIndexesNotStale();
    const tables = sheet.tableIds.map((tableId) => sheet.table(tableId));
    if (tables.some((table) => table.hasQueuedColumnWrites)) {
      throw new Error(
        `Table "${name}" can't be placed on ${sheet.label} while a Table there has column changes queued: flush them first.`,
      );
    }
    const ends = tables.map(
      (table) => table.startColumnIndex + table.columnCount,
    );
    if (ends.length === 0) return undefined;
    return Math.max(...ends);
  }
  private _ensureGridFor(
    placement: CreatedTablePlacement,
    end: { endRowIndex: SheetRowIndex; endColumnIndex: SheetColIndex },
  ): void {
    if (placement.kind === "newTab") {
      this.ss.gatherAddSheetOperation({
        sheetId: placement.sheetGid,
        title: placement.title,
        rowCount: end.endRowIndex,
        columnCount: end.endColumnIndex,
      });
      return;
    }
    const sheet = this.ss.sheet(placement.sheetGid);
    sheet.queueGridRowsThrough(end.endRowIndex);
    sheet.queueGridColumnsThrough(end.endColumnIndex);
  }
  private _fillCells(
    sheetId: number,
    origin: TableOrigin,
    rows: { rowIndex: SheetRowIndex; values: CellValue[] }[],
  ): void {
    for (const { rowIndex, values } of rows) {
      for (const [colIndex, value] of values.entries()) {
        if (value === "") continue;
        this.ss.gatherAddedTableFillCellOperation({
          sheetId,
          rowIndex,
          colIndex: origin.sheetColIndex(colIndex),
          value,
        });
      }
    }
  }
}

export function randomTableId(): string {
  const hexDigitCount = 10;
  let hexDigits = "";
  for (let i = 0; i < hexDigitCount; i++) {
    hexDigits += Math.floor(Math.random() * 16).toString(16);
  }
  return `tbl-${hexDigits}`;
}

function validateSeedRows(
  name: string,
  columnCount: number,
  seedRows: CellValue[][],
): void {
  for (const [rowIndex, values] of seedRows.entries()) {
    if (values.length <= columnCount) continue;
    throw new Error(
      `Table "${name}" can't take seed row ${rowIndex + 1}: it holds ${values.length} values for ${columnCount} columns.`,
    );
  }
}

function createdIds({
  name,
  idPrefix: idPrefixChoice,
  columns,
  tableId = randomTableId(),
}: CreateTableProps): CreatedTable {
  const idPrefix =
    typeof idPrefixChoice === "string"
      ? idPrefixChoice
      : idPrefixes.fromTitle({ title: name, ...idPrefixChoice });
  return {
    tableId,
    idPrefix,
    columnIds: columns.map(
      ({ columnId }) => columnId ?? dimensionIds.col(idPrefix),
    ),
  };
}

function columnTypeProperties(
  columnType: ColumnTypeBaseline | undefined,
): Pick<TableColumnPropertiesAdd, "columnType" | "dataValidationRule"> {
  if (columnType === undefined) return {};
  const { type, rule } = columnType;
  return {
    ...(type === null ? {} : { columnType: type }),
    ...(rule === null ? {} : { dataValidationRule: rule }),
  };
}
