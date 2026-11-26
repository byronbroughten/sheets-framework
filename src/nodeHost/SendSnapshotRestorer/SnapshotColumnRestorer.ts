import type {
  SheetValue,
  TabGrid,
} from "../../00_Source/GoogleSheets/GoogleSheetsAPI/httpSheetsTransport.ts";
import { tableLayout } from "../../01_SpreadsheetSchema/tableLayout.ts";
import type { GoogleCellData } from "../FlushClassifier/flushClassifierTypes.ts";
import type { SnapshotColumn } from "../SendSnapshotTaker.ts";

type GoogleRequest = GoogleAppsScript.Sheets.Schema.Request;

export type ColumnLabel = Pick<SnapshotColumn, "tabTitle" | "header">;

// A saved value whose row is gone, named by its row ID or, on a Table without one, its sheet row.
export interface HomelessRow extends ColumnLabel {
  row: string;
}

// A tab as it stands now, each column from its first row.
export interface CurrentTab {
  grid: TabGrid;
  columns: SheetValue[][];
}

// The ID column before the send and now.
export interface RowIds {
  before: SheetValue[];
  now: SheetValue[];
}

export interface SnapshotColumnRestorerProps {
  column: SnapshotColumn;
  tab: CurrentTab;
  colIndex: number;
  // Undefined where rows match by position: a Table without an ID column, or a tab whose rows didn't move.
  rowIds: RowIds | undefined;
  preSendRowCount: number;
}

export interface ColumnRestoreResult {
  rowCount: number;
  homelessRows: HomelessRow[];
  requests: GoogleRequest[];
}

interface RowTarget {
  rowIndex: number;
  value: SheetValue;
}

// One saved column put back where it stands now: each saved row onto the row it lands on.
export class SnapshotColumnRestorer {
  private column: SnapshotColumn;
  private tab: CurrentTab;
  private colIndex: number;
  private rowIds: RowIds | undefined;
  private preSendRowCount: number;
  constructor(props: SnapshotColumnRestorerProps) {
    this.column = props.column;
    this.tab = props.tab;
    this.colIndex = props.colIndex;
    this.rowIds = props.rowIds;
    this.preSendRowCount = props.preSendRowCount;
  }
  static init(props: SnapshotColumnRestorerProps): SnapshotColumnRestorer {
    return new SnapshotColumnRestorer(props);
  }
  restore(): ColumnRestoreResult {
    const targets: RowTarget[] = [];
    const homelessRows: HomelessRow[] = [];
    const { tabTitle, header } = this.column;
    const dataStart =
      this.column.columnIdRowIndex + tableLayout.headerZoneDepth;
    const landings = this._landings(dataStart);
    for (
      let rowIndex = dataStart;
      rowIndex < this.preSendRowCount;
      rowIndex++
    ) {
      const value = this.column.values[rowIndex] ?? "";
      const { landing, row } = landings(rowIndex);
      if (landing !== undefined) {
        targets.push({ rowIndex: landing, value });
      } else if (value !== "") {
        homelessRows.push({ tabTitle, header, row });
      }
    }
    return {
      rowCount: targets.length,
      homelessRows,
      requests: this._cellWrites(targets),
    };
  }
  // Where each saved row lands now, by row ID or by position, and how a homeless one is named.
  private _landings(
    dataStart: number,
  ): (rowIndex: number) => { landing: number | undefined; row: string } {
    const { rowIds, tab } = this;
    if (rowIds === undefined) {
      return (rowIndex) => ({
        landing: rowIndex < tab.grid.rowCount ? rowIndex : undefined,
        row: `row ${rowIndex + 1}`,
      });
    }
    const rowIndexesNow = rowIndexesById(rowIds.now, dataStart);
    return (rowIndex) => {
      const rowId = rowIds.before[rowIndex] ?? "";
      if (rowId === "") {
        return { landing: undefined, row: `row ${rowIndex + 1}` };
      }
      return {
        landing: rowIndexesNow.get(String(rowId)),
        row: `row ID ${rowId}`,
      };
    };
  }
  // One updateCells per run of consecutive rows, so a row no saved value lands on is left alone.
  private _cellWrites(targets: RowTarget[]): GoogleRequest[] {
    const sorted = targets.toSorted((a, b) => a.rowIndex - b.rowIndex);
    const runs = sorted.reduce((byRun, target) => {
      const run = byRun.at(-1);
      const last = run?.at(-1);
      if (
        run !== undefined &&
        last !== undefined &&
        last.rowIndex + 1 === target.rowIndex
      ) {
        run.push(target);
      } else {
        byRun.push([target]);
      }
      return byRun;
    }, [] as RowTarget[][]);
    return runs.map((run) => ({
      updateCells: {
        range: {
          sheetId: this.column.sheetGid,
          startRowIndex: run[0]?.rowIndex ?? 0,
          endRowIndex: (run.at(-1)?.rowIndex ?? 0) + 1,
          startColumnIndex: this.colIndex,
          endColumnIndex: this.colIndex + 1,
        },
        rows: run.map(({ value }) => ({ values: [enteredCell(value)] })),
        fields: "userEnteredValue",
      },
    }));
  }
}

function rowIndexesById(
  rowIds: SheetValue[],
  dataStart: number,
): Map<string, number> {
  return rowIds.reduce((byId, rowId, rowIndex) => {
    if (rowIndex >= dataStart && rowId !== "") {
      byId.set(String(rowId), rowIndex);
    }
    return byId;
  }, new Map<string, number>());
}

// A value as values:batchGet gave it under FORMULA, entered again; a blank clears the cell.
export function enteredCell(value: SheetValue): GoogleCellData {
  if (value === "") return {};
  if (typeof value === "number") {
    return { userEnteredValue: { numberValue: value } };
  }
  if (typeof value === "boolean") {
    return { userEnteredValue: { boolValue: value } };
  }
  if (value.startsWith("=")) {
    return { userEnteredValue: { formulaValue: value } };
  }
  return { userEnteredValue: { stringValue: value } };
}
