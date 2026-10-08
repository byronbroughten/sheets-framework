import type { TableOrigin } from "../../01_SpreadsheetSchema/TableOrigin";
import {
  buildGridRows,
  type FakeCell,
  type FakeSheetProperties,
} from "../fakeSheetsService";
import { type FakeTablePlacement, fakeTables } from "./fakeTables";

export interface FakeTableColumn {
  columnId: string;
  header: string;
}

export type FakeBodyRow<CN extends string> = Partial<Record<CN, FakeCell>>;

export interface FakeTableSheetProps<
  CN extends string,
> extends FakeTablePlacement {
  sheetId: number;
  title: string;
  name?: string; // The Table's live name; omit to leave it unnamed, as FakeTable does.
  columnConfigs: Record<CN, FakeTableColumn>;
  columnNames: readonly CN[]; // Only these are on the sheet, so a fixture needn't list every configured column.
  bodyRows: readonly FakeBodyRow<CN>[];
}

export const fakeTableSheet = {
  build<CN extends string>({
    sheetId,
    title,
    name,
    columnConfigs,
    columnNames,
    bodyRows,
    ...placement
  }: FakeTableSheetProps<CN>): FakeSheetProperties {
    const origin = fakeTables.origin(placement);
    const columns = columnNames.map((columnName) => columnConfigs[columnName]);
    return {
      sheetId,
      title,
      rows: buildGridRows({
        [origin.headSheetRowIndex("columnId")]: fromFirstColumn(
          origin,
          columns.map((column) => column.columnId),
        ),
        [origin.headerRowIndex]: fromFirstColumn(
          origin,
          columns.map((column) => column.header),
        ),
        ...Object.fromEntries(
          bodyRows.map((row, rowIndex) => [
            origin.sheetRowIndex(rowIndex),
            fromFirstColumn(
              origin,
              columnNames.map((columnName) => row[columnName] ?? null),
            ),
          ]),
        ),
      }),
      tables: [
        {
          name,
          startRowIndex: origin.headerRowIndex,
          startColumnIndex: origin.startColIndex,
          endRowIndex: origin.sheetRowIndex(bodyRows.length),
          endColumnIndex: origin.sheetColIndex(columnNames.length),
        },
      ],
    };
  },
};

function fromFirstColumn(
  origin: TableOrigin,
  cells: readonly FakeCell[],
): FakeCell[] {
  return [...Array<FakeCell>(origin.startColIndex).fill(null), ...cells];
}
