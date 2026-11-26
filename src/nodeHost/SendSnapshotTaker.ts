import type { SheetsHttpTransport } from "../00_Source/GoogleSheets/GoogleSheetsAPI.ts";
import {
  httpColumnValues,
  type SheetValue,
} from "../00_Source/GoogleSheets/GoogleSheetsAPI/httpSheetsTransport.ts";
import { colLetters } from "./colLetters.ts";
import type {
  HeldColumn,
  HeldItem,
} from "./FlushClassifier/flushClassifierTypes.ts";
import {
  columnKey,
  type FlushSnapshotPlan,
  type PlannedColumn,
} from "./flushSnapshotPlan.ts";

// A held column as it stood before the send, head rows included, so a deleted column can be recreated where it was.
export interface SnapshotColumn extends Pick<
  HeldColumn,
  | "sheetGid"
  | "tabTitle"
  | "table"
  | "header"
  | "colIndex"
  | "columnIdRowIndex"
  | "idColIndex"
> {
  columnId: string;
  // From the tab's first row; Google trims the column's trailing blanks.
  values: SheetValue[];
}

// Each row of an affected tab, read from columns the send doesn't change, before and after the send.
export interface TabFingerprints extends Pick<
  HeldColumn,
  "sheetGid" | "tabTitle"
> {
  colIndexes: number[];
  before: SheetValue[][];
  after: SheetValue[][];
}

// What a send changes, as it stood just before the send, and the held items only version history undoes.
export interface SendSnapshot {
  columns: SnapshotColumn[];
  fingerprints: TabFingerprints[];
  versionHistoryItems: HeldItem[];
}

export interface SendSnapshotTakerProps {
  spreadsheetId: string;
  transport: SheetsHttpTransport;
}

type TabFingerprintsBefore = Omit<TabFingerprints, "after">;

// A tab with the rows one read gave its fingerprint columns.
interface TabRows {
  tab: TabFingerprintsBefore;
  rows: SheetValue[][];
}

const fingerprintColumnLimit = 3;

// One snapshot per send: each column is read before the first flush that changes it.
export class SendSnapshotTaker {
  private spreadsheetId: string;
  private transport: SheetsHttpTransport;
  private columns: SnapshotColumn[];
  private capturedColumnKeys: Set<string>;
  private tabs: Map<number, TabFingerprintsBefore>;
  private versionHistoryItems: HeldItem[];
  private hasStarted: boolean;
  constructor(props: SendSnapshotTakerProps) {
    this.spreadsheetId = props.spreadsheetId;
    this.transport = props.transport;
    this.columns = [];
    this.capturedColumnKeys = new Set();
    this.tabs = new Map();
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
    const newTabs = this._newTabs(columns);
    const tabRanges = newTabs.flatMap(fingerprintRanges);
    const ranges = [
      ...freshColumns.map(({ column }) =>
        wholeColumnRange(column.tabTitle, column.colIndex),
      ),
      ...tabRanges,
    ];
    if (ranges.length === 0) return this;
    const values = this._columnValues(ranges);
    for (const [index, planned] of freshColumns.entries()) {
      this._keepColumn(planned, values[index] ?? []);
    }
    const tabValues = values.slice(freshColumns.length);
    for (const { tab, rows } of tabsWithRows(newTabs, tabValues)) {
      this.tabs.set(tab.sheetGid, { ...tab, before: rows });
    }
    return this;
  }
  // Reads the fingerprint columns again after the send's last flush; undefined when the send held nothing.
  finished(): SendSnapshot | undefined {
    if (!this.hasStarted) return undefined;
    const tabs = [...this.tabs.values()];
    const values = this._columnValues(tabs.flatMap(fingerprintRanges));
    const fingerprints = tabsWithRows(tabs, values).map(({ tab, rows }) => ({
      ...tab,
      after: rows,
    }));
    return {
      columns: this.columns,
      fingerprints,
      versionHistoryItems: this.versionHistoryItems,
    };
  }
  // A tab's fingerprint columns are fixed by the first flush that changes it.
  private _newTabs(columns: PlannedColumn[]): TabFingerprintsBefore[] {
    const changedByTab = columns.reduce((byTab, { column }) => {
      if (this.tabs.has(column.sheetGid)) return byTab;
      const changed = byTab.get(column.sheetGid) ?? {
        column,
        colIndexes: new Set<number>(),
      };
      changed.colIndexes.add(column.colIndex);
      byTab.set(column.sheetGid, changed);
      return byTab;
    }, new Map<number, { column: HeldColumn; colIndexes: Set<number> }>());
    return [...changedByTab.values()].map(({ column, colIndexes }) => ({
      sheetGid: column.sheetGid,
      tabTitle: column.tabTitle,
      colIndexes: fingerprintColIndexes(column, colIndexes),
      before: [],
    }));
  }
  private _keepColumn(
    { column, items }: PlannedColumn,
    values: SheetValue[],
  ): void {
    const columnId = values[column.columnIdRowIndex];
    if (typeof columnId !== "string" || columnId === "") {
      this.versionHistoryItems.push(...items);
      return;
    }
    this.capturedColumnKeys.add(columnKey(column));
    const {
      sheetGid,
      tabTitle,
      table,
      header,
      colIndex,
      columnIdRowIndex,
      idColIndex,
    } = column;
    this.columns.push({
      sheetGid,
      tabTitle,
      table,
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
    return httpColumnValues({
      spreadsheetId: this.spreadsheetId,
      transport: this.transport,
      ranges,
    });
  }
}

// The ID column first where it isn't changed, then the tab's first columns not changed.
function fingerprintColIndexes(
  { idColIndex, tabColumnCount }: HeldColumn,
  changed: Set<number>,
): number[] {
  const leading = Array.from({ length: tabColumnCount }, (_, index) => index);
  const candidates =
    idColIndex === undefined ? leading : [idColIndex, ...leading];
  return [...new Set(candidates)]
    .filter((colIndex) => !changed.has(colIndex))
    .slice(0, fingerprintColumnLimit);
}

function fingerprintRanges({
  tabTitle,
  colIndexes,
}: TabFingerprintsBefore): string[] {
  return colIndexes.map((colIndex) => wholeColumnRange(tabTitle, colIndex));
}

function wholeColumnRange(tabTitle: string, colIndex: number): string {
  const letters = colLetters(colIndex);
  return `'${tabTitle.replaceAll("'", "''")}'!${letters}:${letters}`;
}

// The values one read gave each tab's fingerprint columns, in order, paired with each tab as its rows.
function tabsWithRows(
  tabs: TabFingerprintsBefore[],
  values: SheetValue[][],
): TabRows[] {
  const withRows: TabRows[] = [];
  let offset = 0;
  for (const tab of tabs) {
    const end = offset + tab.colIndexes.length;
    withRows.push({ tab, rows: fingerprintRows(values.slice(offset, end)) });
    offset = end;
  }
  return withRows;
}

// The columns turned into rows, a column's trimmed tail padded with "": a tab's fingerprint as the snapshot keeps it.
export function fingerprintRows(columns: SheetValue[][]): SheetValue[][] {
  const rowCount = Math.max(0, ...columns.map((values) => values.length));
  return Array.from({ length: rowCount }, (_, rowIndex) =>
    columns.map((values) => values[rowIndex] ?? ""),
  );
}
