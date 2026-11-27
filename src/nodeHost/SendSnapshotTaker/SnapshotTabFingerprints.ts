import type { SheetValue } from "../../00_Source/GoogleSheets/GoogleSheetsAPI/HttpSpreadsheet.ts";
import { colLetters } from "../colLetters.ts";
import type { HeldColumn } from "../FlushClassifier/flushClassifierTypes.ts";
import type { PlannedColumn } from "../flushSnapshotPlan.ts";
import type {
  TabFingerprints,
  TabFingerprintsBefore,
} from "../SendSnapshotTaker.ts";

const fingerprintColumnLimit = 3;

// Reads the columns of the given ranges, in order, in one request.
export type ColumnReader = (ranges: string[]) => SheetValue[][];

// A tab with the rows one read gave its fingerprint columns.
interface TabRows {
  tab: TabFingerprintsBefore;
  rows: SheetValue[][];
}

// The fingerprints of each tab a send changes, by sheet ID, fixed by the first flush that changes the tab.
export class SnapshotTabFingerprints {
  private tabs: Map<number, TabFingerprintsBefore>;
  constructor() {
    this.tabs = new Map();
  }
  static init(): SnapshotTabFingerprints {
    return new SnapshotTabFingerprints();
  }
  get hasTabs(): boolean {
    return this.tabs.size > 0;
  }
  // One read gives the leading ranges' values, returned, and the fingerprint columns of the tabs this plan is the first to change, kept.
  readNew(
    columns: PlannedColumn[],
    leadingRanges: string[],
    read: ColumnReader,
  ): SheetValue[][] {
    const newTabs = this._newTabs(columns);
    const ranges = [...leadingRanges, ...newTabs.flatMap(fingerprintRanges)];
    if (ranges.length === 0) return [];
    const values = read(ranges);
    const tabValues = values.slice(leadingRanges.length);
    for (const { tab, rows } of tabsWithRows(newTabs, tabValues)) {
      this.tabs.set(tab.sheetGid, { ...tab, before: rows });
    }
    return values.slice(0, leadingRanges.length);
  }
  // A tab's fingerprint columns are fixed by the first flush that changes it.
  // 'after' is compared only with a later read at the same indexes (`SnapshotCurrentTabs`'s `hasMoved`), so a fingerprint column the send changes does no harm.
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
  // Reads the fingerprint columns again, each tab under its title now; a tab with no title now was deleted by the send.
  after(titlesNow: Map<number, string>, read: ColumnReader): TabFingerprints[] {
    const tabs = [...this.tabs.values()];
    const tabsNow = tabs.flatMap((tab) => {
      const tabTitle = titlesNow.get(tab.sheetGid);
      return tabTitle === undefined ? [] : [{ ...tab, tabTitle }];
    });
    const ranges = tabsNow.flatMap(fingerprintRanges);
    const values = ranges.length === 0 ? [] : read(ranges);
    const rowsAfter = new Map(
      tabsWithRows(tabsNow, values).map(({ tab, rows }) => [
        tab.sheetGid,
        rows,
      ]),
    );
    return tabs.map((tab) =>
      fingerprintsAfter(tab, rowsAfter.get(tab.sheetGid)),
    );
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

export function wholeColumnRange(tabTitle: string, colIndex: number): string {
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

// Undefined rows mean the send deleted the tab, so it had no read after.
function fingerprintsAfter(
  tab: TabFingerprintsBefore,
  rows: SheetValue[][] | undefined,
): TabFingerprints {
  if (rows === undefined) return { ...tab, isGoneAfterSend: true };
  return { ...tab, after: rows };
}

// The columns turned into rows, a column's trimmed tail padded with "": a tab's fingerprint as the snapshot keeps it.
export function fingerprintRows(columns: SheetValue[][]): SheetValue[][] {
  const rowCount = Math.max(0, ...columns.map((values) => values.length));
  return Array.from({ length: rowCount }, (_, rowIndex) =>
    columns.map((values) => values[rowIndex] ?? ""),
  );
}
