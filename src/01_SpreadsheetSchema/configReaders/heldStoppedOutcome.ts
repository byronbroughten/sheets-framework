import type { HeldFlushRefusalProps } from "../../00_Source/RawSource/HeldFlushRefusal";
import type { SyncHeldStoppedOutcome, SyncStopCause } from "./syncRunResult";

// A send the gate refused, sync's or a chore's.
export function heldStoppedOutcome({
  kind,
  landedFlushes,
  flushCount,
}: HeldFlushRefusalProps): SyncHeldStoppedOutcome {
  return stoppedOutcome({ landedFlushes, flushCount }, { heldBy: kind });
}

// A send whose batch Google refused whole, after the flushes it counted landed.
export function refusedStoppedOutcome(
  landedFlushes: number,
  googleRefusal: string,
): SyncHeldStoppedOutcome {
  return stoppedOutcome(
    { landedFlushes, flushCount: undefined },
    { googleRefusal },
  );
}

// One stopped at the first flush landed nothing.
function stoppedOutcome(
  { landedFlushes, flushCount }: Omit<HeldFlushRefusalProps, "kind">,
  cause: SyncStopCause,
): SyncHeldStoppedOutcome {
  if (landedFlushes === 0) return { kind: "stoppedBeforeSending", ...cause };
  return { kind: "stoppedPartway", landedFlushes, flushCount, ...cause };
}
