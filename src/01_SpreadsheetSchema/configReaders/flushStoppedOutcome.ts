import type { HeldFlushRefusalProps } from "../../00_Source/RawSource/HeldFlushRefusal";
import type {
  SyncFlushStoppedOutcome,
  SyncLandedFlushes,
  SyncStopCause,
} from "./syncRunResult";

// A send the gate refused, sync's or a chore's.
export function heldStoppedOutcome({
  kind,
  landedFlushes,
  flushCount,
}: HeldFlushRefusalProps): SyncFlushStoppedOutcome {
  return stoppedOutcome({ landedFlushes, flushCount }, { heldBy: kind });
}

// A send whose batch Google refused whole, after the flushes it counted landed.
export function refusedStoppedOutcome(
  landedFlushes: number,
  googleRefusal: string,
): SyncFlushStoppedOutcome {
  return stoppedOutcome({ landedFlushes }, { googleRefusal });
}

// A send whose Fill Row IDs flush threw, after the flushes it counted landed.
export function fillFailedStoppedOutcome(
  landedFlushes: number,
  fillFailure: string,
): SyncFlushStoppedOutcome {
  return stoppedOutcome({ landedFlushes }, { fillFailure });
}

// One stopped at the first flush landed nothing.
function stoppedOutcome(
  { landedFlushes, flushCount }: SyncLandedFlushes,
  cause: SyncStopCause,
): SyncFlushStoppedOutcome {
  if (landedFlushes === 0) return { kind: "stoppedBeforeSending", ...cause };
  return { kind: "stoppedPartway", landedFlushes, flushCount, ...cause };
}
