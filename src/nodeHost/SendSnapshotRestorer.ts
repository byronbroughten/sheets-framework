import {
  HttpSpreadsheet,
  type HttpSpreadsheetProps,
} from "../00_Source/GoogleSheets/GoogleSheetsAPI/HttpSpreadsheet.ts";
import {
  type ColumnLabel,
  type HomelessRow,
} from "./SendSnapshotRestorer/SnapshotColumnRestorer.ts";
import { SnapshotCurrentTabs } from "./SendSnapshotRestorer/SnapshotCurrentTabs.ts";
import type { SendSnapshot, SnapshotColumn } from "./SendSnapshotTaker.ts";

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
      // Columns the send deleted and the restore recreated, with no column type, validation or format of their own.
      recreatedColumns: SnapshotColumn[];
      goneColumns: SnapshotColumn[];
      homelessRows: HomelessRow[];
    };

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
    const tabsNow = SnapshotCurrentTabs.read({
      spreadsheet: this.spreadsheet,
      snapshot: this.snapshot,
    });
    const movedFingerprints = this.snapshot.fingerprints.filter(
      (fingerprints) => tabsNow.hasMoved(fingerprints),
    );
    const changedTabs = movedFingerprints.map(({ tabTitle }) => tabTitle);
    if (changedTabs.length > 0 && !this.isForced) {
      return { kind: "stopped", changedTabs };
    }
    const movedGids = new Set(
      movedFingerprints.map(({ sheetGid }) => sheetGid),
    );
    const { tabs, requests, columns: recreatedColumns } = tabsNow.recreated();
    const restoredColumns: RestoredColumn[] = [];
    const goneColumns: SnapshotColumn[] = [];
    const homelessRows: HomelessRow[] = [];
    for (const column of this.snapshot.columns) {
      const restorer = tabs.columnRestorer(column, movedGids);
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
      recreatedColumns,
      goneColumns,
      homelessRows,
    };
  }
}
