// syncRunResult.ts imports this file back; both imports are type-only, so the cycle is erased at load time.
import type { SyncSubject } from "./syncRunResult";

// A dry run's held items, sync's or a chore's: what the entered-data gate holds until --confirm names the token.
export interface DryRunHeld {
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
  lines: string[];
}
