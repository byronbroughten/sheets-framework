import type { ClaimedTraits } from "../01_SpreadsheetSchema/configReaders/intendedState";
import type { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";

// What the runner knows about this run that the tiers don't.
export interface ChoreRun {
  spreadsheetId: string;
}

// Run from the terminal, as an endpoint is run from the sheet. See docs/architecture/chores.md.
export interface Chore {
  description: string;
  action: (ss: SpreadsheetNamed, run: ChoreRun) => string | void;
  // Only a durable home's chore claims: a `oneOff/` chore is deleted before a claim could matter.
  claimedTraits?: ClaimedTraits;
}
