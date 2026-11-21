import type { HeldFlushRefusalProps } from "../../00_Source/RawSource/HeldFlushRefusal";

// A send stopped at a flush the entered-data gate held, Google refused or, for the fill's own send, failed.
export type FlushStoppedOutcome =
  StoppedBeforeSendingOutcome | StoppedPartwayOutcome;

export type StoppedBeforeSendingOutcome = {
  kind: "stoppedBeforeSending";
} & FlushStopCause;

export type StoppedPartwayOutcome = {
  kind: "stoppedPartway";
} & LandedFlushes &
  FlushStopCause;

export interface LandedFlushes {
  // At least 1 on a partway stop: a send stopped at its first flush is stoppedBeforeSending, since nothing in it landed.
  landedFlushes: number;
  // The dry run's flush count, from its held file; absent when the send had none to read.
  flushCount?: number;
}

export type FlushStopCause = FlushRefusedCause | FillFailedCause;

// The gate held the flush, or Google refused its batch whole, in the reworded text that names the Tables.
export type FlushRefusedCause = { heldBy: HeldBy } | { googleRefusal: string };

// The Fill Row IDs send after the regen threw, by its error's message; the run still writes its files.
export interface FillFailedCause {
  fillFailure: string;
}

export function isFillFailed<OC extends { kind: string }>(
  outcome: OC,
): outcome is OC & FillFailedCause {
  return "fillFailure" in outcome;
}

// "uncovered": the send had no --confirm; "mismatched": the sheet or plan changed since the dry run.
export type HeldBy = "uncovered" | "mismatched";

// A send the gate refused, sync's or a chore's.
export function heldStoppedOutcome({
  kind,
  landedFlushes,
  flushCount,
}: HeldFlushRefusalProps): FlushStoppedOutcome {
  return stoppedOutcome({ landedFlushes, flushCount }, { heldBy: kind });
}

// A send whose batch Google refused whole, after the flushes it counted landed.
export function refusedStoppedOutcome(
  landedFlushes: number,
  googleRefusal: string,
): FlushStoppedOutcome {
  return stoppedOutcome({ landedFlushes }, { googleRefusal });
}

// A send whose Fill Row IDs flush threw, after the flushes it counted landed.
export function fillFailedStoppedOutcome(
  landedFlushes: number,
  fillFailure: string,
): FlushStoppedOutcome {
  return stoppedOutcome({ landedFlushes }, { fillFailure });
}

// One stopped at the first flush landed nothing.
function stoppedOutcome(
  { landedFlushes, flushCount }: LandedFlushes,
  cause: FlushStopCause,
): FlushStoppedOutcome {
  if (landedFlushes === 0) return { kind: "stoppedBeforeSending", ...cause };
  return { kind: "stoppedPartway", landedFlushes, flushCount, ...cause };
}
