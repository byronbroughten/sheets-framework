import {
  HttpSpreadsheet,
  type HttpSpreadsheetProps,
  type SheetValue,
} from "../00_Source/GoogleSheets/GoogleSheetsAPI/HttpSpreadsheet.ts";
import type {
  HeldColumn,
  HeldItem,
} from "./FlushClassifier/flushClassifierTypes.ts";
import {
  columnKey,
  type FlushSnapshotPlan,
  type PlannedColumn,
} from "./flushSnapshotPlan.ts";
import {
  SnapshotTabFingerprints,
  wholeColumnRange,
} from "./SendSnapshotTaker/SnapshotTabFingerprints.ts";

// A held column as it stood before the send, head rows included, so a deleted column can be recreated where it was.
export interface SnapshotColumn extends Pick<
  HeldColumn,
  | "sheetGid"
  | "tabTitle"
  | "table"
  | "tableId"
  | "header"
  | "colIndex"
  | "columnIdRowIndex"
  | "idColIndex"
> {
  columnId: string;
  // From the tab's first row; Google trims the column's trailing blanks.
  values: SheetValue[];
}

// Each row of an affected tab, read from columns the send doesn't change, before the send.
export interface TabFingerprintsBefore extends Pick<
  HeldColumn,
  "sheetGid" | "tabTitle"
> {
  colIndexes: number[];
  before: SheetValue[][];
}

// The same columns read again after the send.
interface TabFingerprintsAfter extends TabFingerprintsBefore {
  after: SheetValue[][];
}

// A tab the send deleted, so it has no rows after it.
interface TabGoneAfterSend extends TabFingerprintsBefore {
  isGoneAfterSend: true;
}

export type TabFingerprints = TabFingerprintsAfter | TabGoneAfterSend;

// What a send changes, as it stood just before the send, and the held items only version history undoes.
export interface SendSnapshot {
  columns: SnapshotColumn[];
  fingerprints: TabFingerprints[];
  versionHistoryItems: HeldItem[];
}

export type SendSnapshotTakerProps = HttpSpreadsheetProps;

// One snapshot per send: each column is read before the first flush that changes it.
export class SendSnapshotTaker {
  private spreadsheet: HttpSpreadsheet;
  private columns: SnapshotColumn[];
  private capturedColumnKeys: Set<string>;
  private tabFingerprints: SnapshotTabFingerprints;
  private versionHistoryItems: HeldItem[];
  private hasStarted: boolean;
  constructor(props: SendSnapshotTakerProps) {
    this.spreadsheet = HttpSpreadsheet.init(props);
    this.columns = [];
    this.capturedColumnKeys = new Set();
    this.tabFingerprints = SnapshotTabFingerprints.init();
    this.versionHistoryItems = [];
    this.hasStarted = false;
  }
  static init(props: SendSnapshotTakerProps): SendSnapshotTaker {
    return new SendSnapshotTaker(props);
  }
  // Called just before a covered flush goes out; a plan holding nothing takes no read.
  capture({ columns, versionHistoryItems }: FlushSnapshotPlan): this {
    if (columns.length === 0 && versionHistoryItems.length === 0) return this;
    this.hasStarted = true;
    this.versionHistoryItems.push(...versionHistoryItems);
    const freshColumns = columns.filter(
      ({ column }) => !this.capturedColumnKeys.has(columnKey(column)),
    );
    const freshRanges = freshColumns.map(({ column }) =>
      wholeColumnRange(column.tabTitle, column.colIndex),
    );
    const freshValues = this.tabFingerprints.readNew(
      columns,
      freshRanges,
      (ranges) => this._columnValues(ranges),
    );
    for (const [index, planned] of freshColumns.entries()) {
      this._keepColumn(planned, freshValues[index] ?? []);
    }
    return this;
  }
  // Reads the fingerprint columns again after the send's last flush, each tab under its title then; undefined when the send held nothing.
  finished(): SendSnapshot | undefined {
    if (!this.hasStarted) return undefined;
    const titlesNow = this._tabTitlesNow();
    const fingerprints = this.tabFingerprints.after(titlesNow, (ranges) =>
      this._columnValues(ranges),
    );
    return {
      columns: this.columns,
      fingerprints,
      versionHistoryItems: this.versionHistoryItems,
    };
  }
  // By sheet ID, since the send may have renamed or deleted a tab; a send that changed no tab's columns takes no read.
  private _tabTitlesNow(): Map<number, string> {
    if (!this.tabFingerprints.hasTabs) return new Map();
    return new Map(
      this.spreadsheet
        .tabGrids()
        .map(({ sheetGid, title }) => [sheetGid, title]),
    );
  }
  private _keepColumn(
    { column, items }: PlannedColumn,
    values: SheetValue[],
  ): void {
    const columnId = values[column.columnIdRowIndex];
    // RunReads.heldColumn drops a column whose ID read blank, so this guards an ID cleared after the dry run, or one the run never read.
    if (typeof columnId !== "string" || columnId === "") {
      this.versionHistoryItems.push(...items);
      return;
    }
    this.capturedColumnKeys.add(columnKey(column));
    const {
      sheetGid,
      tabTitle,
      table,
      tableId,
      header,
      colIndex,
      columnIdRowIndex,
      idColIndex,
    } = column;
    this.columns.push({
      sheetGid,
      tabTitle,
      table,
      tableId,
      header,
      colIndex,
      columnIdRowIndex,
      idColIndex,
      columnId,
      values,
    });
  }
  private _columnValues(ranges: string[]): SheetValue[][] {
    if (ranges.length === 0) return [];
    return this.spreadsheet.columnValues(ranges);
  }
}
