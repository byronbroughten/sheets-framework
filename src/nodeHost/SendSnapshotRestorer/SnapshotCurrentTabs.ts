import {
  type HttpSpreadsheet,
  type SheetValue,
  type TabGrid,
} from "../../00_Source/GoogleSheets/GoogleSheetsAPI/HttpSpreadsheet.ts";
import { colLetters } from "../colLetters.ts";
import {
  type SendSnapshot,
  type SnapshotColumn,
  type TabFingerprints,
} from "../SendSnapshotTaker.ts";
import { fingerprintRows } from "../SendSnapshotTaker/SnapshotTabFingerprints.ts";
import { SnapshotColumnRecreator } from "./SnapshotColumnRecreator.ts";
import {
  type CurrentTab,
  type RowIds,
  SnapshotColumnRestorer,
  tableOf,
} from "./SnapshotColumnRestorer.ts";

type GoogleRequest = GoogleAppsScript.Sheets.Schema.Request;

export interface SnapshotCurrentTabsProps {
  snapshot: SendSnapshot;
  tabs: Map<number, CurrentTab>;
}

export interface SnapshotCurrentTabsReadProps {
  spreadsheet: HttpSpreadsheet;
  snapshot: SendSnapshot;
}

// The tabs as the recreated columns leave them, and the requests that recreate them.
interface TabsRecreation {
  tabs: SnapshotCurrentTabs;
  requests: GoogleRequest[];
  columns: SnapshotColumn[];
}

// The snapshot's tabs as they stand now, by sheet ID. Recreating deleted columns returns a new one, so a moved-check made on the tabs as read can't see the recreated ones.
export class SnapshotCurrentTabs {
  private snapshot: SendSnapshot;
  private tabs: Map<number, CurrentTab>;
  constructor(props: SnapshotCurrentTabsProps) {
    this.snapshot = props.snapshot;
    this.tabs = props.tabs;
  }
  static init(props: SnapshotCurrentTabsProps): SnapshotCurrentTabs {
    return new SnapshotCurrentTabs(props);
  }
  // One GET for the tabs' sizes and Tables, then one values:batchGet of each affected tab whole.
  static read({
    spreadsheet,
    snapshot,
  }: SnapshotCurrentTabsReadProps): SnapshotCurrentTabs {
    const { columns, fingerprints } = snapshot;
    const sheetGids = new Set(
      [...columns, ...fingerprints].map(({ sheetGid }) => sheetGid),
    );
    const grids = spreadsheet
      .tabGrids()
      .filter(
        (grid) =>
          sheetGids.has(grid.sheetGid) &&
          grid.rowCount > 0 &&
          grid.columnCount > 0,
      );
    if (grids.length === 0) {
      return SnapshotCurrentTabs.init({ snapshot, tabs: new Map() });
    }
    const values = spreadsheet.rangeColumns(grids.map(tabRange));
    const tabs = new Map(
      grids.map((grid, index): [number, CurrentTab] => [
        grid.sheetGid,
        { grid, columns: values[index] ?? [] },
      ]),
    );
    return SnapshotCurrentTabs.init({ snapshot, tabs });
  }
  // A tab whose fingerprint columns differ from just after the send: rows were added, removed or reordered.
  hasMoved(fingerprints: TabFingerprints): boolean {
    const tab = this.tabs.get(fingerprints.sheetGid);
    // A tab the send deleted, still gone, leaves the others to restore.
    if ("isGoneAfterSend" in fingerprints) return tab !== undefined;
    if (tab === undefined) return true;
    const now = fingerprintRows(
      fingerprints.colIndexes.map((colIndex) => tab.columns[colIndex] ?? []),
    );
    return JSON.stringify(now) !== JSON.stringify(fingerprints.after);
  }
  // A deleted column comes back first, so its values are written as an overwritten column's are.
  recreated(): TabsRecreation {
    const recreated = new Map(this.tabs);
    const requests: GoogleRequest[] = [];
    const columns: SnapshotColumn[] = [];
    for (const [sheetGid, tab] of this.tabs) {
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
      columns.push(...recreation.columns);
    }
    return {
      tabs: SnapshotCurrentTabs.init({
        snapshot: this.snapshot,
        tabs: recreated,
      }),
      requests,
      columns,
    };
  }
  // Undefined where the column's tab or Table, found by its recorded ID, is gone, so it couldn't be recreated.
  columnRestorer(
    column: SnapshotColumn,
    movedGids: Set<number>,
  ): SnapshotColumnRestorer | undefined {
    const tab = this.tabs.get(column.sheetGid);
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
