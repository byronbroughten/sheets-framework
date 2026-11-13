import type { SyncPlannerResult } from "./SyncPlanner";
import type {
  SyncDryRunOutcome,
  SyncDryRunResult,
  SyncRefusal,
  SyncRunCommon,
  SyncRunFrame,
  SyncRunResult,
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
  ran({ planned, notes, typedTraitChanges }: SyncRunPlanned): SyncRunResult {
    const common = this._common({
      refusals: [],
      planned,
      notes,
      typedTraitChanges,
    });
    if (planned.isStopped) return this._unsent(common, { kind: "stopped" });
    if (!this.isSending) return dryRunResult(common, { kind: "ready" });
    return {
      ...sendResult(common, {
        kind: "sent",
        summary: sentSummary(planned.plan.length),
      }),
      landed: planned.plan.map(({ subject, what }) => ({
        kind: "send",
        subject,
        what,
      })),
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

function sentSummary(landedCount: number): string {
  if (landedCount === 0) return "nothing to send";
  if (landedCount === 1) return "1 alteration landed";
  return `${landedCount} alterations landed`;
}
