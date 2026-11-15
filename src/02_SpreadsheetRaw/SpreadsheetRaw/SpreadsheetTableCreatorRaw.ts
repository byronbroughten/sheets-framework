import type { CellValue } from "../../00_Source/CellValues/cellValues";
import {
  type SheetColIndex,
  SheetIndex,
  type SheetRowIndex,
} from "../../00_Source/RawSource/SheetIndex";
import { dimensionIds } from "../../01_SpreadsheetSchema/dimensionIds";
import { idPrefixes } from "../../01_SpreadsheetSchema/idPrefixes";
import { TableOrigin } from "../../01_SpreadsheetSchema/TableOrigin";
import { SpreadsheetBaseRaw } from "../ClassBases/SpreadsheetBaseRaw";
import { queuedColumnType } from "../ClassTypes/queuedColumnType";
import type {
  CreatedTable,
  CreatedTablePlacement,
  CreateTableProps,
} from "../ClassTypes/StateRaw";
import { SpreadsheetRaw } from "../SpreadsheetRaw";

// A created Table holds one empty body row unless it is seeded.
const minBodyRowCount = 1;

interface TableSize {
  bodyRowCount: number;
  columnCount: number;
}

type BesideTablesPlacement = Extract<
  CreatedTablePlacement,
  { kind: "besideTables" }
>;

export class SpreadsheetTableCreatorRaw extends SpreadsheetBaseRaw {
  get ss(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
  // IDs are settled before any queue, so a dry run's withheld flush and the caller's record carry the same ones.
  createTable(props: CreateTableProps): CreatedTable {
    const { name, placement, columns, seedRows = [] } = props;
    validateSeedRows(name, columns.length, seedRows);
    const created = createdIds(props);
    const size = {
      bodyRowCount: Math.max(minBodyRowCount, seedRows.length),
      columnCount: columns.length,
    };
    const origin = this._queueGridFor(name, placement, size);
    const { endRowIndex, endColumnIndex } = tableEnd(origin, size);
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
        ...(columnType === undefined
          ? {}
          : queuedColumnType.properties(columnType)),
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
  // Returns where the Table starts, having queued a tab or grid that holds it.
  private _queueGridFor(
    name: string,
    placement: CreatedTablePlacement,
    size: TableSize,
  ): TableOrigin {
    if (placement.kind === "newTab") {
      const origin = TableOrigin.expected();
      const end = tableEnd(origin, size);
      this.ss.gatherAddSheetOperation({
        sheetId: placement.sheetGid,
        title: placement.title,
        rowCount: end.endRowIndex,
        columnCount: end.endColumnIndex,
      });
      return origin;
    }
    const origin = this._originBeside(name, placement);
    const end = tableEnd(origin, size);
    const sheet = this.ss.sheet(placement.sheetGid);
    sheet.queueGridRowsThrough(end.endRowIndex);
    sheet.queueGridColumnsThrough(end.endColumnIndex);
    return origin;
  }
  // Only a full properties fetch lists every Table on the tab; a grid fetch lists only those its ranges overlap.
  private _originBeside(
    name: string,
    { sheetGid, gap = 0 }: BesideTablesPlacement,
  ): TableOrigin {
    this.ss.ensureAllSheetPropertiesAreFetched();
    if (!this.ss.gidIsActive(sheetGid)) {
      throw new Error(
        `Table "${name}" can't be placed beside the Tables on GID ${sheetGid}: no fetched tab has it.`,
      );
    }
    const sheet = this.ss.sheet(sheetGid);
    sheet.assertRowIndexesNotStale();
    if (sheet.hasQueuedTableColumnWrites) {
      throw new Error(
        `Table "${name}" can't be placed on ${sheet.label} while a Table there has column changes queued: flush them first.`,
      );
    }
    const expected = TableOrigin.expected();
    const rightmostEnd = sheet.rightmostTableEndColIndex;
    if (rightmostEnd === undefined) return expected;
    return new TableOrigin({
      headerRowIndex: expected.headerRowIndex,
      startColIndex: SheetIndex.col(rightmostEnd + gap),
    });
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

function tableEnd(
  origin: TableOrigin,
  { bodyRowCount, columnCount }: TableSize,
): { endRowIndex: SheetRowIndex; endColumnIndex: SheetColIndex } {
  return {
    endRowIndex: origin.sheetRowIndex(bodyRowCount),
    endColumnIndex: origin.sheetColIndex(columnCount),
  };
}
