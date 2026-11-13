import type { SyncPlannerResult } from "./SyncPlanner";
import type {
  SyncDryRunOutcome,
  SyncDryRunResult,
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

export interface SyncRunPlanned {
  planned: SyncPlannerResult;
  notes: string[];
  typedTraitChanges: SyncTypedTraitChange[];
  // The sent traits the refetch shows Sheets kept another value for; they leave the landed list.
  didntLand: SyncSendLine[];
}

interface SyncRunCommonProps {
  refusals: SyncRefusal[];
  planned?: SyncPlannerResult;
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
  }: SyncRunPlanned): SyncRunResult {
    const common = this._common({
      refusals: [],
      planned,
      notes,
      typedTraitChanges,
    });
    if (planned.isStopped) return this._unsent(common, { kind: "stopped" });
    if (!this.isSending) return dryRunResult(common, { kind: "ready" });
    const didntLandKeys = new Set(
      didntLand.map(({ subject }) => JSON.stringify(subject)),
    );
    const landed = planned.plan.filter(
      ({ subject }) => !didntLandKeys.has(JSON.stringify(subject)),
    );
    return {
      ...sendResult(common, {
        kind: "sent",
        summary: sentSummary(planned.plan.length, didntLand.length),
      }),
      landed: landed.map(({ subject, what }) => ({
        kind: "send",
        subject,
        what,
      })),
      didntLand,
    };
  }
  private _common({
    refusals,
    planned,
    notes,
    typedTraitChanges,
  }: SyncRunCommonProps): SyncRunCommon {
    return {
      ...this.frame,
      checks: [],
      refusals,
      stops: planned?.stops ?? [],
      plan: planned?.plan ?? [],
      forcedTraits: planned?.forcedTraits ?? [],
      notices: [],
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

function sentSummary(sentCount: number, didntLandCount: number): string {
  if (sentCount === 0) return "nothing to send";
  const landedCount = sentCount - didntLandCount;
  const landed =
    landedCount === 0
      ? "nothing landed"
      : `${landedCount} ${landedCount === 1 ? "alteration" : "alterations"} landed`;
  if (didntLandCount === 0) return landed;
  return `${landed}; ${didntLandCount} didn't land`;
}
