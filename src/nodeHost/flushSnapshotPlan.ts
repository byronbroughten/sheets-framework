import type { ForcedTrait } from "../01_SpreadsheetSchema/configReaders/syncRunResult.ts";
import type {
  HeldCellItem,
  HeldColumn,
  HeldItem,
} from "./FlushClassifier/flushClassifierTypes.ts";
import type { RunReads } from "./FlushClassifier/RunReads.ts";
import { isHeldCellItem } from "./heldGroups.ts";

export interface PlannedColumn {
  column: HeldColumn;
  items: HeldCellItem[];
}

// What a send snapshot reads before one flush goes out: each held column, and the held items it can't hold.
export interface FlushSnapshotPlan {
  columns: PlannedColumn[];
  versionHistoryItems: HeldItem[];
}

export interface FlushSnapshotPlanProps {
  reads: RunReads;
  items: HeldItem[];
  // Forced traits name no cell; the first flush's plan names them once for the send.
  forcedTraits: ForcedTrait[];
  // An earlier flush judged against the same reads shifted cells, so a read column's index may name another live column: every held cell goes to version history.
  hasShiftedCells: boolean;
}

export function flushSnapshotPlan({
  reads,
  items,
  forcedTraits,
  hasShiftedCells,
}: FlushSnapshotPlanProps): FlushSnapshotPlan {
  const columns = new Map<string, PlannedColumn>();
  const versionHistoryItems: HeldItem[] = forcedTraits.map((trait) => ({
    kind: "forcedTrait",
    trait,
  }));
  for (const item of items) {
    if (!isHeldCellItem(item)) {
      versionHistoryItems.push(item);
      continue;
    }
    const column = hasShiftedCells ? undefined : reads.heldColumn(item.cell);
    if (column === undefined) {
      versionHistoryItems.push(item);
      continue;
    }
    const key = columnKey(column);
    const planned = columns.get(key) ?? { column, items: [] };
    planned.items.push(item);
    columns.set(key, planned);
  }
  return { columns: [...columns.values()], versionHistoryItems };
}

export function columnKey({ sheetGid, colIndex }: HeldColumn): string {
  return `${sheetGid}:${colIndex}`;
}
