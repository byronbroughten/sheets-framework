// syncRunResult.ts imports this file back; both imports are type-only, so the cycle is erased at load time.
import type { SyncSubject } from "./syncRunResult.ts";

// A dry run's held block and forced traits, sync's or a chore's.
export interface DryRunHeld {
  // Undefined when no flush held an item, so no --confirm is needed.
  block: HeldBlock | undefined;
  // What --force sends over the live value, which needs no --confirm; a chore forces none.
  forcedTraits: HeldGroup[];
}

// What the entered-data gate holds until --confirm names the token.
export interface HeldBlock {
  groups: HeldGroup[];
  token: string;
  fullListPath: string;
}

export interface HeldGroup {
  // Absent for a held item no Table or tab names, such as a find-and-replace.
  subject?: SyncSubject;
  what: string;
  // Entered-data cells this group holds; absent for a held item that isn't cells, such as a dropdown reset.
  cells?: number;
  // True for an item the send snapshot can't hold, by flushSnapshotPlan's rule.
  isUndoneOnlyByVersionHistory?: boolean;
  lines: string[];
}
