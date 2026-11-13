// Uncovered: no token was given; mismatched: a token was, and it doesn't cover this flush's held items.
export type HeldFlushRefusalKind = "uncovered" | "mismatched";

export interface HeldFlushRefusalProps {
  kind: HeldFlushRefusalKind;
  landedFlushes: number;
  // The flushes the given held list counted; undefined when no list was given.
  flushCount: number | undefined;
}

// A host held a flush that could lose entered data (ADR 0034); the flushes before it have landed.
export class HeldFlushRefusal extends Error {
  readonly kind: HeldFlushRefusalKind;
  readonly landedFlushes: number;
  readonly flushCount: number | undefined;
  constructor({ kind, landedFlushes, flushCount }: HeldFlushRefusalProps) {
    super(
      `Held flush ${landedFlushes + 1} (${kind}): it could lose entered data, and no token covers its held list. ${landedFlushes} earlier flushes landed.`,
    );
    this.kind = kind;
    this.landedFlushes = landedFlushes;
    this.flushCount = flushCount;
  }
}
