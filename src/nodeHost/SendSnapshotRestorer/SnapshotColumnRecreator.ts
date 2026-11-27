import type { SnapshotColumn } from "../SendSnapshotTaker.ts";
import { RecreatedColumn } from "./RecreatedColumn.ts";
import { type CurrentTab, tableOf } from "./SnapshotColumnRestorer.ts";

type GoogleRequest = GoogleAppsScript.Sheets.Schema.Request;

export interface SnapshotColumnRecreatorProps {
  tab: CurrentTab;
  // The tab's saved columns whose column ID is gone from it.
  columns: SnapshotColumn[];
}

// The tab as the recreations leave it, so a later column-ID lookup finds each recreated column.
export interface ColumnRecreation {
  tab: CurrentTab;
  requests: GoogleRequest[];
}

// Each deleted column back at its saved position, head rows and column ID included, as the framework inserts one: from the column ID row to the Table's last row.
export class SnapshotColumnRecreator {
  private tab: CurrentTab;
  private columns: SnapshotColumn[];
  constructor(props: SnapshotColumnRecreatorProps) {
    this.tab = props.tab;
    this.columns = props.columns;
  }
  static init(props: SnapshotColumnRecreatorProps): SnapshotColumnRecreator {
    return new SnapshotColumnRecreator(props);
  }
  // Leftmost first, so each saved position is read in the pre-send layout the ones before it rebuild.
  recreate(): ColumnRecreation {
    const sorted = this.columns.toSorted((a, b) => a.colIndex - b.colIndex);
    return sorted.reduce<ColumnRecreation>(
      (recreation, column) => recreatedOn(recreation, column),
      { tab: this.tab, requests: [] },
    );
  }
}

// A column whose Table's ID is gone is left as it is, for the restore to list as not restored.
function recreatedOn(
  { tab, requests }: ColumnRecreation,
  column: SnapshotColumn,
): ColumnRecreation {
  const table = tableOf(tab, column);
  if (table === undefined) return { tab, requests };
  const recreated = RecreatedColumn.init({ column, table });
  return {
    tab: recreated.tabAfter(tab),
    requests: [...requests, ...recreated.requests],
  };
}
