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
}

export function flushSnapshotPlan({
  reads,
  items,
  forcedTraits,
}: FlushSnapshotPlanProps): FlushSnapshotPlan {
  const columns = new Map<string, PlannedColumn>();
  const versionHistoryItems: HeldItem[] = forcedTraits.map((trait) => ({
    kind: "forcedTrait",
    trait,
  }));
  for (const item of items) {
    const column = isHeldCellItem(item) && reads.heldColumn(item.cell);
    if (!isHeldCellItem(item) || !column) {
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
