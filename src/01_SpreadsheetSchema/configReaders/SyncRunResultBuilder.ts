import type { HeldFlushRefusal } from "../../00_Source/RawSource/HeldFlushRefusal";
import { hasLanded } from "./didntLand";
import type { SyncPlannerResult } from "./SyncPlanner";
import type {
  SyncDryRunOutcome,
  SyncDryRunResult,
  SyncNotice,
  SyncPlanItem,
  SyncRefusal,
  SyncRunCommon,
  SyncRunFrame,
  SyncRunResult,
  SyncSendLine,
  SyncSendOutcome,
  SyncSendResult,
  SyncTypedTraitChange,
} from "./syncRunResult";

export interface SyncRunResultBuilderProps {
  isSending: boolean;
  frame: SyncRunFrame;
}

// sync's own sends, which no alteration asked for, and their notices: Fill Row IDs' checkbox and fill.
export interface SyncRunDerived {
  plan: SyncPlanItem[];
  notices: SyncNotice[];
}

export interface SyncRunPlanned {
  planned: SyncPlannerResult;
  notes: string[];
  typedTraitChanges: SyncTypedTraitChange[];
  // The sent traits the refetch shows Sheets kept another value for; they leave the landed list.
  didntLand: SyncSendLine[];
  derived: SyncRunDerived;
}

interface SyncRunCommonProps {
  refusals: SyncRefusal[];
  planned?: SyncPlannerResult;
  derived?: SyncRunDerived;
  notes: string[];
  typedTraitChanges: SyncTypedTraitChange[];
}

// Builds one run's SyncRunResult from its mode and frame; the fields later tickets fill stay empty.
export class SyncRunResultBuilder {
  readonly isSending: boolean;
  readonly frame: SyncRunFrame;
  constructor({ isSending, frame }: SyncRunResultBuilderProps) {
    this.isSending = isSending;
    this.frame = frame;
  }
  static init(props: SyncRunResultBuilderProps): SyncRunResultBuilder {
    return new SyncRunResultBuilder(props);
  }
  refused(refusals: SyncRefusal[]): SyncRunResult {
    const common = this._common({ refusals, notes: [], typedTraitChanges: [] });
    return this._unsent(common, { kind: "refused" });
  }
  ran({
    planned,
    notes,
    typedTraitChanges,
    didntLand,
    derived,
  }: SyncRunPlanned): SyncRunResult {
    const common = this._common({
      refusals: [],
      planned,
      derived,
      notes,
      typedTraitChanges,
    });
    if (planned.isStopped) return this._unsent(common, { kind: "stopped" });
    if (!this.isSending) return dryRunResult(common, { kind: "ready" });
    const isLanded = hasLanded(didntLand);
    const landed = [
      ...planned.plan.filter(({ subject }) => isLanded(subject)),
      ...derived.plan,
    ];
    return {
      ...sendResult(common, {
        kind: "sent",
        summary: sentSummary({
          sentCount: planned.plan.length,
          didntLandCount: didntLand.length,
          derivedCount: derived.plan.length,
        }),
      }),
      landed: landed.map(({ subject, what }) => ({
        kind: "send",
        subject,
        what,
      })),
      didntLand,
    };
  }
  // A send the host refused at a held flush; nothing after it was sent.
  heldStopped(
    planned: SyncPlannerResult,
    { kind, landedFlushes, flushCount }: HeldFlushRefusal,
    derived: SyncRunDerived,
  ): SyncRunResult {
    const common = this._common({
      refusals: [],
      planned,
      derived,
      notes: [],
      typedTraitChanges: [],
    });
    if (landedFlushes === 0) {
      return sendResult(common, { kind: "stoppedBeforeSending", heldBy: kind });
    }
    return sendResult(common, {
      kind: "stoppedPartway",
      landedFlushes,
      flushCount,
      heldBy: kind,
    });
  }
  private _common({
    refusals,
    planned,
    derived,
    notes,
    typedTraitChanges,
  }: SyncRunCommonProps): SyncRunCommon {
    return {
      ...this.frame,
      checks: [],
      refusals,
      stops: planned?.stops ?? [],
      plan: [...(planned?.plan ?? []), ...(derived?.plan ?? [])],
      forcedTraits: planned?.forcedTraits ?? [],
      notices: derived?.notices ?? [],
      configs: {
        heldBack: planned?.configs.heldBack ?? [],
        renamed: [],
        namesMoved: [],
        rewrites: [],
        typedTraitChanges,
        notes,
      },
    };
  }
  // A refused or stopped run sends nothing, in either mode.
  private _unsent(
    common: SyncRunCommon,
    outcome: { kind: "refused" } | { kind: "stopped" },
  ): SyncRunResult {
    if (this.isSending) return sendResult(common, outcome);
    return dryRunResult(common, outcome);
  }
}

interface SentCounts {
  sentCount: number;
  didntLandCount: number;
  // sync's own sends, counted apart from the alterations.
  derivedCount: number;
}

function dryRunResult(
  common: SyncRunCommon,
  outcome: SyncDryRunOutcome,
): SyncDryRunResult {
  return { ...common, mode: "dry run", outcome };
}

function sendResult(
  common: SyncRunCommon,
  outcome: SyncSendOutcome,
): SyncSendResult {
  return {
    ...common,
    mode: "send",
    outcome,
    landed: [],
    notSent: [],
    didntLand: [],
  };
}

function sentSummary({
  sentCount,
  didntLandCount,
  derivedCount,
}: SentCounts): string {
  const derived = `${derivedCount} Fill Row IDs ${derivedCount === 1 ? "send" : "sends"}`;
  if (sentCount === 0) {
    if (derivedCount === 0) return "nothing to send";
    return derived;
  }
  const alterations = alterationsSummary(sentCount, didntLandCount);
  if (derivedCount === 0) return alterations;
  return `${alterations}; ${derived}`;
}

function alterationsSummary(sentCount: number, didntLandCount: number): string {
  const landedCount = sentCount - didntLandCount;
  const landed =
    landedCount === 0
      ? "nothing landed"
      : `${landedCount} ${landedCount === 1 ? "alteration" : "alterations"} landed`;
  if (didntLandCount === 0) return landed;
  return `${landed}; ${didntLandCount} didn't land`;
}
