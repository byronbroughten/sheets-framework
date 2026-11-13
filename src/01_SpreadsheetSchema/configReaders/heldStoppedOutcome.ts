import type { HeldFlushRefusalProps } from "../../00_Source/RawSource/HeldFlushRefusal";
import type { SyncHeldStoppedOutcome } from "./syncRunResult";

// A send the gate refused, sync's or a chore's; one refused at the first flush landed nothing.
export function heldStoppedOutcome({
  kind,
  landedFlushes,
  flushCount,
}: HeldFlushRefusalProps): SyncHeldStoppedOutcome {
  if (landedFlushes === 0) {
    return { kind: "stoppedBeforeSending", heldBy: kind };
  }
  return { kind: "stoppedPartway", landedFlushes, flushCount, heldBy: kind };
}
