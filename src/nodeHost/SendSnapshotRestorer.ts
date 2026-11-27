import {
  HttpSpreadsheet,
  type HttpSpreadsheetProps,
  type SheetValue,
  type TabGrid,
} from "../00_Source/GoogleSheets/GoogleSheetsAPI/HttpSpreadsheet.ts";
import { colLetters } from "./colLetters.ts";
import { SnapshotColumnRecreator } from "./SendSnapshotRestorer/SnapshotColumnRecreator.ts";
import {
  type ColumnLabel,
  type CurrentTab,
  type HomelessRow,
  type RowIds,
  SnapshotColumnRestorer,
  tableOf,
} from "./SendSnapshotRestorer/SnapshotColumnRestorer.ts";
import {
  fingerprintRows,
  type SendSnapshot,
  type SnapshotColumn,
  type TabFingerprints,
} from "./SendSnapshotTaker.ts";

type GoogleRequest = GoogleAppsScript.Sheets.Schema.Request;

export interface SendSnapshotRestorerProps extends HttpSpreadsheetProps {
  snapshot: SendSnapshot;
  isForced: boolean;
}

export interface RestoredColumn extends ColumnLabel {
  rowCount: number;
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

// The tabs as the recreated columns leave them, and the requests that recreate them.
interface TabsRecreation {
  tabs: Map<number, CurrentTab>;
  requests: GoogleRequest[];
}

// Puts a send's snapshot back: each saved column by its column ID, a deleted one recreated first, its rows by position on a tab whose rows didn't move, else by row ID where the Table has an ID column.
export class SendSnapshotRestorer {
  private spreadsheet: HttpSpreadsheet;
  private snapshot: SendSnapshot;
  private isForced: boolean;
  constructor(props: SendSnapshotRestorerProps) {
    this.spreadsheet = HttpSpreadsheet.init(props);
    this.snapshot = props.snapshot;
    this.isForced = props.isForced;
  }
  static init(props: SendSnapshotRestorerProps): SendSnapshotRestorer {
    return new SendSnapshotRestorer(props);
  }
  // Writes nothing when a tab's rows moved since the send, unless forced.
  restore(): SnapshotRestore {
    const tabsNow = this._currentTabs();
    const movedFingerprints = this.snapshot.fingerprints.filter(
      (fingerprints) => hasMoved(fingerprints, tabsNow),
    );
    const changedTabs = movedFingerprints.map(({ tabTitle }) => tabTitle);
    if (changedTabs.length > 0 && !this.isForced) {
      return { kind: "stopped", changedTabs };
    }
    const movedGids = new Set(
      movedFingerprints.map(({ sheetGid }) => sheetGid),
    );
    const { tabs, requests } = this._recreatedTabs(tabsNow);
    const restoredColumns: RestoredColumn[] = [];
    const goneColumns: SnapshotColumn[] = [];
    const homelessRows: HomelessRow[] = [];
    for (const column of this.snapshot.columns) {
      const restorer = this._columnRestorer(column, tabs, movedGids);
      if (restorer === undefined) {
        goneColumns.push(column);
        continue;
      }
      const restored = restorer.restore();
      homelessRows.push(...restored.homelessRows);
      requests.push(...restored.requests);
      const { tabTitle, header } = column;
      restoredColumns.push({ tabTitle, header, rowCount: restored.rowCount });
    }
    if (requests.length > 0) {
      // Past the entered-data gate: a restore undoes a send the gate already held.
      this.spreadsheet.batchUpdate(requests);
    }
    return {
      kind: "restored",
      changedTabs,
      restoredColumns,
      goneColumns,
      homelessRows,
    };
  }
  // A deleted column comes back first, so its values are written as an overwritten column's are.
  private _recreatedTabs(tabs: Map<number, CurrentTab>): TabsRecreation {
    const recreated = new Map(tabs);
    const requests: GoogleRequest[] = [];
    for (const [sheetGid, tab] of tabs) {
      const deleted = this.snapshot.columns.filter(
        (column) =>
          column.sheetGid === sheetGid &&
          colIndexOf(tab, column.columnIdRowIndex, column.columnId) ===
            undefined,
      );
      if (deleted.length === 0) continue;
      const recreation = SnapshotColumnRecreator.init({
        tab,
        columns: deleted,
      }).recreate();
      recreated.set(sheetGid, recreation.tab);
      requests.push(...recreation.requests);
    }
    return { tabs: recreated, requests };
  }
  // One GET for the tabs' sizes and Tables, then one values:batchGet of each affected tab whole.
  private _currentTabs(): Map<number, CurrentTab> {
    const { columns, fingerprints } = this.snapshot;
    const sheetGids = new Set(
      [...columns, ...fingerprints].map(({ sheetGid }) => sheetGid),
    );
    const grids = this.spreadsheet
      .tabGrids()
      .filter(
        (grid) =>
          sheetGids.has(grid.sheetGid) &&
          grid.rowCount > 0 &&
          grid.columnCount > 0,
      );
    if (grids.length === 0) return new Map();
    const values = this.spreadsheet.rangeColumns(grids.map(tabRange));
    return new Map(
      grids.map((grid, index) => [
        grid.sheetGid,
        { grid, columns: values[index] ?? [] },
      ]),
    );
  }
  // Undefined where the column's tab or Table, found by its recorded ID, is gone, so it couldn't be recreated.
  private _columnRestorer(
    column: SnapshotColumn,
    tabs: Map<number, CurrentTab>,
    movedGids: Set<number>,
  ): SnapshotColumnRestorer | undefined {
    const tab = tabs.get(column.sheetGid);
    if (tab === undefined) return undefined;
    const table = tableOf(tab, column);
    const colIndex = colIndexOf(tab, column.columnIdRowIndex, column.columnId);
    if (table === undefined || colIndex === undefined) return undefined;
    const fingerprints = this._fingerprintsOf(column.sheetGid);
    // Fingerprints that still match mean no row moved, so position is exact, even for a row whose ID the send filled.
    const isUnmoved =
      fingerprints !== undefined && !movedGids.has(column.sheetGid);
    return SnapshotColumnRestorer.init({
      column,
      table,
      colIndex,
      rowIds: isUnmoved ? undefined : this._rowIds(column, tab),
      preSendRowCount: Math.max(
        column.values.length,
        fingerprints?.before.length ?? 0,
      ),
    });
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
  fingerprints: TabFingerprints,
  tabs: Map<number, CurrentTab>,
): boolean {
  const tab = tabs.get(fingerprints.sheetGid);
  // A tab the send deleted, still gone, leaves the others to restore.
  if ("isGoneAfterSend" in fingerprints) return tab !== undefined;
  if (tab === undefined) return true;
  const now = fingerprintRows(
    fingerprints.colIndexes.map((colIndex) => tab.columns[colIndex] ?? []),
  );
  return JSON.stringify(now) !== JSON.stringify(fingerprints.after);
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
