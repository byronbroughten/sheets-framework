// Uncovered: no token was given; mismatched: a token was, and it doesn't cover this flush's held items.
export type HeldFlushRefusalKind = "uncovered" | "mismatched";

// A host held a flush that could lose entered data (ADR 0034); the flushes before it have landed.
export class HeldFlushRefusal extends Error {
  constructor(
    readonly kind: HeldFlushRefusalKind,
    readonly landedFlushes: number,
  ) {
    super(
      `Held flush ${landedFlushes + 1} (${kind}): it could lose entered data, and no token covers its held list. ${landedFlushes} earlier flushes landed.`,
    );
  }
}
