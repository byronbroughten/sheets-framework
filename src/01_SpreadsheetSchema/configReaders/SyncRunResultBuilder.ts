import type { HeldFlushRefusal } from "../../00_Source/RawSource/HeldFlushRefusal";
import { hasLanded } from "./didntLand";
import { heldStoppedOutcome } from "./heldStoppedOutcome";
import type { SyncPlannerResult } from "./SyncPlanner";
import {
  fillRowIdsName,
  type SyncDryRunOutcome,
  type SyncDryRunResult,
  type SyncNotice,
  type SyncPlanItem,
  type SyncRefusal,
  type SyncRunCommon,
  type SyncRunFrame,
  type SyncRunResult,
  type SyncSendLine,
  type SyncSendOutcome,
  type SyncSendResult,
  type SyncTypedTraitChange,
} from "./syncRunResult";

export interface SyncRunResultBuilderProps {
  isSending: boolean;
  frame: SyncRunFrame;
}

// Fill Row IDs' checkbox and fill: sync's own sends, which no alteration asked for, and their notices.
export interface SyncRunRowIds {
  plan: SyncPlanItem[];
  notices: SyncNotice[];
}

export interface SyncRunHeldStopped {
  planned: SyncPlannerResult;
  refusal: HeldFlushRefusal;
  rowIds: SyncRunRowIds;
}

export interface SyncRunPlanned {
  planned: SyncPlannerResult;
  notes: string[];
  typedTraitChanges: SyncTypedTraitChange[];
  // The sent traits the refetch shows Sheets kept another value for; they leave the landed list.
  didntLand: SyncSendLine[];
  rowIds: SyncRunRowIds;
}

interface SyncRunCommonProps {
  refusals: SyncRefusal[];
  planned?: SyncPlannerResult;
  rowIds?: SyncRunRowIds;
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
    rowIds,
  }: SyncRunPlanned): SyncRunResult {
    const common = this._common({
      refusals: [],
      planned,
      rowIds,
      notes,
      typedTraitChanges,
    });
    if (planned.isStopped) return this._unsent(common, { kind: "stopped" });
    if (!this.isSending) return dryRunResult(common, { kind: "ready" });
    const isLanded = hasLanded(didntLand);
    const landed = [
      ...planned.plan.filter(({ subject }) => isLanded(subject)),
      ...rowIds.plan,
    ];
    return {
      ...sendResult(common, {
        kind: "sent",
        summary: sentSummary({
          sentCount: planned.plan.length,
          didntLandCount: didntLand.length,
          rowIdsSendCount: rowIds.plan.length,
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
  heldStopped({ planned, refusal, rowIds }: SyncRunHeldStopped): SyncRunResult {
    const common = this._common({
      refusals: [],
      planned,
      rowIds,
      notes: [],
      typedTraitChanges: [],
    });
    return sendResult(common, heldStoppedOutcome(refusal));
  }
  private _common({
    refusals,
    planned,
    rowIds,
    notes,
    typedTraitChanges,
  }: SyncRunCommonProps): SyncRunCommon {
    return {
      ...this.frame,
      checks: [],
      refusals,
      stops: planned?.stops ?? [],
      plan: [...(planned?.plan ?? []), ...(rowIds?.plan ?? [])],
      forcedTraits: planned?.forcedTraits ?? [],
      notices: rowIds?.notices ?? [],
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
  rowIdsSendCount: number;
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
  rowIdsSendCount,
}: SentCounts): string {
  const rowIdsSends = `${rowIdsSendCount} ${fillRowIdsName} ${rowIdsSendCount === 1 ? "send" : "sends"}`;
  if (sentCount === 0) {
    if (rowIdsSendCount === 0) return "nothing to send";
    return rowIdsSends;
  }
  const alterations = alterationsSummary(sentCount, didntLandCount);
  if (rowIdsSendCount === 0) return alterations;
  return `${alterations}; ${rowIdsSends}`;
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
