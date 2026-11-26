import type { SheetsHttpTransport } from "../00_Source/GoogleSheets/GoogleSheetsAPI.ts";
import {
  httpBatchUpdate,
  httpRangeColumns,
  httpTabGrids,
  type SheetValue,
  type TabGrid,
} from "../00_Source/GoogleSheets/GoogleSheetsAPI/httpSheetsTransport.ts";
import { tableLayout } from "../01_SpreadsheetSchema/tableLayout.ts";
import { colLetters } from "./colLetters.ts";
import type { GoogleCellData } from "./FlushClassifier/flushClassifierTypes.ts";
import {
  fingerprintRows,
  type SendSnapshot,
  type SnapshotColumn,
  type TabFingerprints,
} from "./SendSnapshotTaker.ts";

type GoogleRequest = GoogleAppsScript.Sheets.Schema.Request;

export interface SendSnapshotRestorerProps {
  spreadsheetId: string;
  transport: SheetsHttpTransport;
  snapshot: SendSnapshot;
  isForced: boolean;
}

type ColumnLabel = Pick<SnapshotColumn, "tabTitle" | "header">;

export interface RestoredColumn extends ColumnLabel {
  rowCount: number;
}

// A saved value whose row is gone, named by its row ID or, on a Table without one, its sheet row.
export interface HomelessRow extends ColumnLabel {
  row: string;
}

export type SnapshotRestore =
  | { kind: "stopped"; changedTabs: string[] }
  | {
      kind: "restored";
      // Restored over anyway, by --force.
      changedTabs: string[];
      restoredColumns: RestoredColumn[];
      goneColumns: SnapshotColumn[];
      homelessRows: HomelessRow[];
    };

// A tab as it stands now, each column from its first row.
interface CurrentTab {
  grid: TabGrid;
  columns: SheetValue[][];
}

interface RowTarget {
  rowIndex: number;
  value: SheetValue;
}

interface ColumnRestore {
  column: SnapshotColumn;
  tab: CurrentTab;
  colIndex: number;
  // Undefined on a Table without an ID column, whose rows match by position.
  rowIds: RowIds | undefined;
  preSendRowCount: number;
}

interface RowIds {
  before: SheetValue[];
  now: SheetValue[];
}

interface ColumnTargets {
  targets: RowTarget[];
  homelessRows: HomelessRow[];
}

// Puts a send's snapshot back: each saved column by its column ID, its rows by row ID where the Table has an ID column and by position otherwise.
export class SendSnapshotRestorer {
  private spreadsheetId: string;
  private transport: SheetsHttpTransport;
  private snapshot: SendSnapshot;
  private isForced: boolean;
  constructor(props: SendSnapshotRestorerProps) {
    this.spreadsheetId = props.spreadsheetId;
    this.transport = props.transport;
    this.snapshot = props.snapshot;
    this.isForced = props.isForced;
  }
  static init(props: SendSnapshotRestorerProps): SendSnapshotRestorer {
    return new SendSnapshotRestorer(props);
  }
  // Writes nothing when a tab's rows moved since the send, unless forced.
  restore(): SnapshotRestore {
    const tabs = this._currentTabs();
    const changedTabs = this.snapshot.fingerprints
      .filter((fingerprints) => hasMoved(fingerprints, tabs))
      .map(({ tabTitle }) => tabTitle);
    if (changedTabs.length > 0 && !this.isForced) {
      return { kind: "stopped", changedTabs };
    }
    const restoredColumns: RestoredColumn[] = [];
    const goneColumns: SnapshotColumn[] = [];
    const homelessRows: HomelessRow[] = [];
    const requests: GoogleRequest[] = [];
    for (const column of this.snapshot.columns) {
      const columnRestore = this._columnRestore(column, tabs);
      if (columnRestore === undefined) {
        goneColumns.push(column);
        continue;
      }
      const targets = columnTargets(columnRestore);
      homelessRows.push(...targets.homelessRows);
      requests.push(...cellWrites(columnRestore, targets.targets));
      const { tabTitle, header } = column;
      restoredColumns.push({
        tabTitle,
        header,
        rowCount: targets.targets.length,
      });
    }
    if (requests.length > 0) {
      httpBatchUpdate({
        spreadsheetId: this.spreadsheetId,
        transport: this.transport,
        requests,
      });
    }
    return {
      kind: "restored",
      changedTabs,
      restoredColumns,
      goneColumns,
      homelessRows,
    };
  }
  // One GET for the tabs' sizes, then one values:batchGet of each affected tab whole.
  private _currentTabs(): Map<number, CurrentTab> {
    const { columns, fingerprints } = this.snapshot;
    const sheetGids = new Set(
      [...columns, ...fingerprints].map(({ sheetGid }) => sheetGid),
    );
    const call = {
      spreadsheetId: this.spreadsheetId,
      transport: this.transport,
    };
    const grids = httpTabGrids(call).filter(
      (grid) =>
        sheetGids.has(grid.sheetGid) &&
        grid.rowCount > 0 &&
        grid.columnCount > 0,
    );
    if (grids.length === 0) return new Map();
    const values = httpRangeColumns({ ...call, ranges: grids.map(tabRange) });
    return new Map(
      grids.map((grid, index) => [
        grid.sheetGid,
        { grid, columns: values[index] ?? [] },
      ]),
    );
  }
  // Undefined where the column's tab or column ID is gone.
  private _columnRestore(
    column: SnapshotColumn,
    tabs: Map<number, CurrentTab>,
  ): ColumnRestore | undefined {
    const tab = tabs.get(column.sheetGid);
    if (tab === undefined) return undefined;
    const colIndex = colIndexOf(tab, column.columnIdRowIndex, column.columnId);
    if (colIndex === undefined) return undefined;
    const fingerprints = this._fingerprintsOf(column.sheetGid);
    return {
      column,
      tab,
      colIndex,
      rowIds: this._rowIds(column, tab),
      preSendRowCount: Math.max(
        column.values.length,
        fingerprints?.before.length ?? 0,
      ),
    };
  }
  private _fingerprintsOf(sheetGid: number): TabFingerprints | undefined {
    return this.snapshot.fingerprints.find(
      (fingerprints) => fingerprints.sheetGid === sheetGid,
    );
  }
  // The ID column before the send and now, found now by the column ID it had then.
  private _rowIds(column: SnapshotColumn, tab: CurrentTab): RowIds | undefined {
    const { idColIndex, columnIdRowIndex } = column;
    if (idColIndex === undefined) return undefined;
    const before = this._idColumnBefore(column.sheetGid, idColIndex);
    const idColumnId = before[columnIdRowIndex];
    const nowColIndex =
      typeof idColumnId === "string" && idColumnId !== ""
        ? colIndexOf(tab, columnIdRowIndex, idColumnId)
        : undefined;
    const now =
      nowColIndex === undefined ? [] : (tab.columns[nowColIndex] ?? []);
    return { before, now };
  }
  // A saved ID column, else the fingerprint column it is whenever the send left it alone.
  private _idColumnBefore(sheetGid: number, idColIndex: number): SheetValue[] {
    const saved = this.snapshot.columns.find(
      (column) =>
        column.sheetGid === sheetGid && column.colIndex === idColIndex,
    );
    if (saved !== undefined) return saved.values;
    const fingerprints = this._fingerprintsOf(sheetGid);
    const index = fingerprints?.colIndexes.indexOf(idColIndex) ?? -1;
    // An ID column the send changed but couldn't hold: no row can be matched.
    if (fingerprints === undefined || index === -1) return [];
    return fingerprints.before.map((row) => row[index] ?? "");
  }
}

function tabRange({ title, rowCount, columnCount }: TabGrid): string {
  const lastColumn = colLetters(columnCount - 1);
  return `'${title.replaceAll("'", "''")}'!A1:${lastColumn}${rowCount}`;
}

// A tab whose fingerprint columns differ from just after the send: rows were added, removed or reordered.
function hasMoved(
  { sheetGid, colIndexes, after }: TabFingerprints,
  tabs: Map<number, CurrentTab>,
): boolean {
  const tab = tabs.get(sheetGid);
  if (tab === undefined) return true;
  const now = fingerprintRows(
    colIndexes.map((colIndex) => tab.columns[colIndex] ?? []),
  );
  return JSON.stringify(now) !== JSON.stringify(after);
}

function colIndexOf(
  tab: CurrentTab,
  columnIdRowIndex: number,
  columnId: string,
): number | undefined {
  const colIndex = tab.columns.findIndex(
    (values) => values[columnIdRowIndex] === columnId,
  );
  return colIndex === -1 ? undefined : colIndex;
}

// Each saved data row and the row it lands on now; a nonblank one with nowhere to land is homeless.
function columnTargets({
  column,
  tab,
  rowIds,
  preSendRowCount,
}: ColumnRestore): ColumnTargets {
  const dataStart = column.columnIdRowIndex + tableLayout.headerZoneDepth;
  const rowIndexesNow =
    rowIds === undefined ? undefined : rowIndexesById(rowIds.now, dataStart);
  const targets: RowTarget[] = [];
  const homelessRows: HomelessRow[] = [];
  for (let rowIndex = dataStart; rowIndex < preSendRowCount; rowIndex++) {
    const value = column.values[rowIndex] ?? "";
    let landing: number | undefined = rowIndex;
    let row = `row ${rowIndex + 1}`;
    if (rowIds !== undefined && rowIndexesNow !== undefined) {
      const rowId = rowIds.before[rowIndex] ?? "";
      landing = rowId === "" ? undefined : rowIndexesNow.get(String(rowId));
      if (rowId !== "") row = `row ID ${rowId}`;
    } else if (rowIndex >= tab.grid.rowCount) {
      landing = undefined;
    }
    if (landing !== undefined) {
      targets.push({ rowIndex: landing, value });
    } else if (value !== "") {
      homelessRows.push({
        tabTitle: column.tabTitle,
        header: column.header,
        row,
      });
    }
  }
  return { targets, homelessRows };
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

// One updateCells per run of consecutive rows, so a row no saved value lands on is left alone.
function cellWrites(
  { column, colIndex }: ColumnRestore,
  targets: RowTarget[],
): GoogleRequest[] {
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
        sheetId: column.sheetGid,
        startRowIndex: run[0]?.rowIndex ?? 0,
        endRowIndex: (run.at(-1)?.rowIndex ?? 0) + 1,
        startColumnIndex: colIndex,
        endColumnIndex: colIndex + 1,
      },
      rows: run.map(({ value }) => ({ values: [enteredCell(value)] })),
      fields: "userEnteredValue",
    },
  }));
}

// A value as values:batchGet gave it under FORMULA, entered again; a blank clears the cell.
function enteredCell(value: SheetValue): GoogleCellData {
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
