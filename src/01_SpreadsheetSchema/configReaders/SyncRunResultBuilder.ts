import { hasLanded } from "./didntLand";
import type { FlushStoppedOutcome } from "./flushStoppedOutcome";
import type { SyncNameRewriter } from "./SyncNameRewriter";
import {
  newColumnLines,
  newTableLines,
  type SyncPlannerResult,
} from "./SyncPlanner";
import {
  fillRowIdsName,
  heldBackChangeWhat,
  type SyncCreatedColumn,
  type SyncCreatedTable,
  type SyncDryRunOutcome,
  type SyncDryRunResult,
  type SyncHeldBackChange,
  type SyncLandedItem,
  type SyncNotice,
  type SyncPlanItem,
  type SyncRefusal,
  type SyncRunCommon,
  type SyncRunFrame,
  type SyncRunResult,
  type SyncSendLine,
  type SyncSendOutcome,
  type SyncSendResult,
  type SyncSubject,
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

export interface SyncRunFlushStopped {
  planned: SyncPlannerResult;
  outcome: FlushStoppedOutcome;
  rowIds: SyncRunRowIds;
}

// A send that stopped after its flushes landed in part, with the Row ID plan lines that never landed: Google's refusal of a rule flush, or the Fill Row IDs send that threw after the regen.
export interface SyncRunSendStop {
  outcome: FlushStoppedOutcome;
  unlanded: SyncPlanItem[];
}

export interface SyncRunPlanned {
  planned: SyncPlannerResult;
  notes: string[];
  typedTraitChanges: SyncTypedTraitChange[];
  // The sent traits the refetch shows Sheets kept another value for; they leave the landed list.
  didntLand: SyncSendLine[];
  // The new columns the refetch found, in place of their plan lines on the landed list.
  createdColumns: SyncCreatedColumn[];
  // The new Tables the refetch found, in place of their plan lines on the landed list.
  createdTables: SyncCreatedTable[];
  rowIds: SyncRunRowIds;
  nameRewriter: SyncNameRewriter;
  // Left out of the plan, so a send lists each as not sent and a dry run as a notice.
  heldBackChanges: SyncHeldBackChange[];
  sendStop?: SyncRunSendStop;
}

interface SyncRunCommonProps {
  refusals: SyncRefusal[];
  planned?: SyncPlannerResult;
  rowIds?: SyncRunRowIds;
  notes: string[];
  typedTraitChanges: SyncTypedTraitChange[];
  // Absent on a run that writes no files, so nothing it names was rewritten.
  nameRewriter?: SyncNameRewriter;
}

// Builds one run's SyncRunResult from its mode and frame.
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
    createdColumns,
    createdTables,
    rowIds,
    nameRewriter,
    heldBackChanges,
    sendStop,
  }: SyncRunPlanned): SyncRunResult {
    const common = this._common({
      refusals: [],
      planned,
      rowIds,
      notes,
      typedTraitChanges,
      nameRewriter,
    });
    if (planned.isStopped) return this._unsent(common, { kind: "stopped" });
    if (!this.isSending) {
      const notices: SyncNotice[] = [
        ...common.notices,
        ...heldBackChanges.map((rule) => ({
          kind: "ruleHeldBack" as const,
          ...rule,
        })),
      ];
      return dryRunResult({ ...common, notices }, { kind: "ready" });
    }
    const isLanded = hasLanded(didntLand);
    const landedAlterations: SyncLandedItem[] = [
      ...planned.plan
        .filter(
          // Each new column or Table lands as its created line, carrying the minted IDs its plan line couldn't.
          ({ subject }) =>
            isLanded(subject) &&
            !newColumnLines.isSubject(subject) &&
            !newTableLines.isSubject(subject),
        )
        .map(landedSend),
      ...createdColumns,
      ...createdTables,
    ];
    const metOrLandedCount =
      landedAlterations.length + planned.alreadyMet.length;
    const outcome: SyncSendOutcome = sendStop?.outcome ?? {
      kind: "sent",
      summary: sentSummary({
        sentCount: planned.plan.length,
        didntLandCount: didntLand.length,
        rowIdsSendCount: rowIds.plan.length,
        alreadyMetCount: planned.alreadyMet.length,
      }),
    };
    const unlanded = new Set(sendStop?.unlanded);
    const landedRowIds = rowIds.plan.filter((item) => !unlanded.has(item));
    return {
      ...sendResult(common, outcome),
      landed: [...landedAlterations, ...landedRowIds.map(landedSend)],
      notSent: heldBackChanges.map((rule) =>
        asRewritten(heldBackLine(rule), nameRewriter),
      ),
      didntLand: didntLand.map((line) => asRewritten(line, nameRewriter)),
      hasEmptiedAlterations:
        nameRewriter.isAlterationsEmpty && metOrLandedCount > 0,
    };
  }
  // A send the host held or Google refused at a flush; nothing after it was sent.
  flushStopped({
    planned,
    outcome,
    rowIds,
  }: SyncRunFlushStopped): SyncRunResult {
    const common = this._common({
      refusals: [],
      planned,
      rowIds,
      notes: [],
      typedTraitChanges: [],
    });
    return sendResult(common, outcome);
  }
  private _common({
    refusals,
    planned,
    rowIds,
    notes,
    typedTraitChanges,
    nameRewriter,
  }: SyncRunCommonProps): SyncRunCommon {
    return {
      ...this.frame,
      checks: planned?.checks ?? [],
      refusals,
      stops: (planned?.stops ?? []).map((stop) => {
        if (stop.kind !== "conflict") return stop;
        return asRewritten(stop, nameRewriter);
      }),
      plan: [
        ...(planned?.plan ?? []).map((item) => asRewritten(item, nameRewriter)),
        ...(rowIds?.plan ?? []),
      ],
      alreadyMet: planned?.alreadyMet ?? [],
      forcedTraits: planned?.forcedTraits ?? [],
      notices: rowIds?.notices ?? [],
      configs: {
        heldBack: (planned?.configs.heldBack ?? []).map((trait) =>
          asRewritten(trait, nameRewriter),
        ),
        renamed: nameRewriter?.renamed ?? [],
        namesMoved: nameRewriter?.namesMoved ?? [],
        rewrites: nameRewriter?.rewrites ?? [],
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
  alreadyMetCount: number;
}

// A pending entry's line under the name sync rewrote it to, once the regen renamed what it names.
function asRewritten<IT extends { subject: SyncSubject }>(
  item: IT,
  nameRewriter: SyncNameRewriter | undefined,
): IT {
  if (nameRewriter === undefined) return item;
  return { ...item, subject: nameRewriter.asRewritten(item.subject) };
}

function heldBackLine(rule: SyncHeldBackChange): SyncSendLine {
  return { subject: rule.subject, what: heldBackChangeWhat(rule) };
}

function landedSend({ subject, what }: SyncSendLine): SyncLandedItem {
  return { kind: "send", subject, what };
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
    hasEmptiedAlterations: false,
  };
}

function sentSummary({ alreadyMetCount, ...counts }: SentCounts): string {
  const sends = landedAndRowIdsSummary(counts);
  if (alreadyMetCount === 0) return sends;
  return `${sends}; ${alterationCount(alreadyMetCount)} already met`;
}

function landedAndRowIdsSummary({
  sentCount,
  didntLandCount,
  rowIdsSendCount,
}: Omit<SentCounts, "alreadyMetCount">): string {
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
      : `${alterationCount(landedCount)} landed`;
  if (didntLandCount === 0) return landed;
  return `${landed}; ${didntLandCount} didn't land`;
}

function alterationCount(count: number): string {
  return `${count} ${count === 1 ? "alteration" : "alterations"}`;
}
