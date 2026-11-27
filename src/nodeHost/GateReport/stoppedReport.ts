import {
  type FillFailedCause,
  type FlushRefusedCause,
  type FlushStoppedOutcome,
  type HeldBy,
  isFillFailed,
  type LandedFlushes,
} from "../../01_SpreadsheetSchema/configReaders/flushStoppedOutcome.ts";
import { fillRowIdsName } from "../../01_SpreadsheetSchema/configReaders/syncRunResult.ts";

type FillFailedOutcome = FlushStoppedOutcome & FillFailedCause;

export const stoppedReport = {
  // The result line of a send the gate or Google refused, sync's or a chore's, or whose fill failed.
  flushStoppedLine(outcome: FlushStoppedOutcome): string {
    if (isFillFailed(outcome)) return fillFailedLine(outcome);
    const refused = refusedWords(outcome);
    if (outcome.kind === "stoppedBeforeSending") {
      return `STOPPED: the send was refused before any flush landed${refused} Nothing was sent.`;
    }
    return stoppedPartwayLine(
      outcome,
      ` was refused${refused} Nothing after it was sent.`,
    );
  },
};

// The fill is the send's last flush, after the files are built.
function fillFailedLine(outcome: FillFailedOutcome): string {
  const failed = `the ${fillRowIdsName} send, failed: "${outcome.fillFailure}". The blank IDs stay blank; the configs were still written.`;
  if (outcome.kind === "stoppedBeforeSending") {
    return `STOPPED: no flush landed; flush 1, ${failed}`;
  }
  return stoppedPartwayLine(outcome, `, ${failed}`);
}

// Google's reworded refusal ends in its own full stop.
function refusedWords(cause: FlushRefusedCause): string {
  if ("heldBy" in cause) return `, because ${heldByReason[cause.heldBy]}.`;
  return `: Google refused its batch and applied none of it. ${cause.googleRefusal}`;
}

const heldByReason: Record<HeldBy, string> = {
  uncovered: "it holds items and the send had no --confirm",
  mismatched:
    "the sheet or the plan changed after the dry run, so --confirm doesn't cover its held items",
};

// The stopping flush is the one after the last that landed; its cause's words follow it.
function stoppedPartwayLine(
  { landedFlushes, flushCount }: LandedFlushes,
  stoppedWords: string,
): string {
  const landed = flushSpan(landedFlushes, flushCount);
  return `STOPPED PARTWAY: ${landed}; flush ${landedFlushes + 1}${stoppedWords}`;
}

function flushSpan(
  landedFlushes: number,
  flushCount: number | undefined,
): string {
  const ofCount = flushCount === undefined ? "" : ` of ${flushCount}`;
  if (landedFlushes === 1) return `flush 1${ofCount} landed`;
  return `flushes 1 to ${landedFlushes}${ofCount} landed`;
}
