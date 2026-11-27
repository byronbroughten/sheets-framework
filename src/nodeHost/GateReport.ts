import type {
  DryRunHeld,
  HeldBlock,
  HeldGroup,
} from "../01_SpreadsheetSchema/configReaders/dryRunHeld.ts";
import {
  type FillFailedCause,
  type FlushRefusedCause,
  type FlushStoppedOutcome,
  type HeldBy,
  isFillFailed,
  type LandedFlushes,
} from "../01_SpreadsheetSchema/configReaders/flushStoppedOutcome.ts";
import { fillRowIdsName } from "../01_SpreadsheetSchema/configReaders/syncRunResult.ts";
import type { HeldItem } from "./FlushClassifier/flushClassifierTypes.ts";
import { heldGroups } from "./heldGroups.ts";
import { ReportText } from "./ReportText.ts";

const { section, indent, formatCount, counted, subjectLabel } = ReportText;

export const gateSectionTitles = {
  held: "Held",
  forced: "Forced: sent over the live value, no --confirm needed",
} as const;

type FillFailedOutcome = FlushStoppedOutcome & FillFailedCause;

// The entered-data gate's report pieces, which sync's report and a chore's output both print.
export const GateReport = {
  // A dry run's forced traits, then its held items, their full list's path and the token covering them; a chore's dry run prints it too.
  heldSections({ block, forcedGroups }: DryRunHeld): string[] {
    return [
      ...section(
        gateSectionTitles.forced,
        forcedGroups.flatMap(heldGroupLines),
      ),
      ...(block === undefined ? [] : heldBlockLines(block)),
    ];
  },
  // The held block's items only; a forced trait needs no --confirm, so it never counts.
  heldCount({ block }: DryRunHeld): number {
    if (block === undefined) return 0;
    return blockCount(block);
  },
  // The result line of a send the gate or Google refused, sync's or a chore's, or whose fill failed.
  flushStoppedLine(outcome: FlushStoppedOutcome): string {
    if (isFillFailed(outcome)) return fillFailedLine(outcome);
    const refused = refusedWords(outcome);
    if (outcome.kind === "stoppedBeforeSending") {
      return `STOPPED: the send was refused before any flush landed${refused} Nothing was sent.`;
    }
    return stoppedPartwayLine(
      outcome,
      ` was refused${refused} Nothing after it was sent.`,
    );
  },
  rerunDryRunLine(command: string): string {
    return `Rerun the dry run: ${command}`;
  },
  // Beside a send's undo line: the held items the send snapshot couldn't hold; nothing when it held them all.
  versionHistoryLines(items: HeldItem[]): string[] {
    const groups = heldGroups({
      flushes: [items],
      versionHistoryItems: [items],
    });
    if (groups.length === 0) return [];
    return [
      `${versionHistoryWords}:`,
      ...indent(
        2,
        groups.flatMap((group) => groupBlock(group, [])),
      ),
    ];
  },
};

function heldBlockLines(block: HeldBlock): string[] {
  return section(heldTitle(block), [
    ...block.groups.flatMap(heldGroupLines),
    `Full list, every cell with its value: ${block.fullListPath}`,
    `Token ${block.token}: binds this target and exactly this list.`,
    `Rerun with --send --confirm ${block.token} to send these.`,
  ]);
}

function heldTitle(block: HeldBlock): string {
  const confirm = "sent only with --confirm";
  const cellCounts = block.groups.flatMap(({ cells }) =>
    cells === undefined ? [] : [cells],
  );
  const count = blockCount(block);
  if (cellCounts.length < block.groups.length) {
    return `${gateSectionTitles.held}: ${counted(count, "item", "items")}, ${confirm}`;
  }
  let sum = "";
  if (cellCounts.length > 1) {
    sum = `: ${cellCounts.map(formatCount).join(" + ")}`;
  }
  return `${gateSectionTitles.held}: ${counted(count, "cell", "cells")}${sum}, ${confirm}`;
}

function blockCount({ groups }: HeldBlock): number {
  return groups.reduce((count, group) => count + (group.cells ?? 1), 0);
}

function heldGroupLines(group: HeldGroup): string[] {
  const undo = group.isUndoneOnlyByVersionHistory ? [versionHistoryLine] : [];
  return groupBlock(group, undo);
}

// A group's heading, then its extra lines and its own, indented.
function groupBlock(
  { subject, what, lines }: HeldGroup,
  extraLines: string[],
): string[] {
  const heading =
    subject === undefined ? what : `${subjectLabel(subject)}  ${what}`;
  return [heading, ...indent(2, [...extraLines, ...lines])];
}

const versionHistoryWords = "undone only by the spreadsheet's version history";

const versionHistoryLine = `${versionHistoryWords}: the send snapshot can't hold it`;

// The fill is the send's last flush, after the files are built.
function fillFailedLine(outcome: FillFailedOutcome): string {
  const failed = `the ${fillRowIdsName} send, failed: "${outcome.fillFailure}". The blank IDs stay blank; the configs were still written.`;
  if (outcome.kind === "stoppedBeforeSending") {
    return `STOPPED: no flush landed; flush 1, ${failed}`;
  }
  return stoppedPartwayLine(outcome, `, ${failed}`);
}

// Google's reworded refusal ends in its own full stop.
function refusedWords(cause: FlushRefusedCause): string {
  if ("heldBy" in cause) return `, because ${heldByReason[cause.heldBy]}.`;
  return `: Google refused its batch and applied none of it. ${cause.googleRefusal}`;
}

const heldByReason: Record<HeldBy, string> = {
  uncovered: "it holds items and the send had no --confirm",
  mismatched:
    "the sheet or the plan changed after the dry run, so --confirm doesn't cover its held items",
};

// The stopping flush is the one after the last that landed; its cause's words follow it.
function stoppedPartwayLine(
  { landedFlushes, flushCount }: LandedFlushes,
  stoppedWords: string,
): string {
  const landed = flushSpan(landedFlushes, flushCount);
  return `STOPPED PARTWAY: ${landed}; flush ${landedFlushes + 1}${stoppedWords}`;
}

function flushSpan(
  landedFlushes: number,
  flushCount: number | undefined,
): string {
  const ofCount = flushCount === undefined ? "" : ` of ${flushCount}`;
  if (landedFlushes === 1) return `flush 1${ofCount} landed`;
  return `flushes 1 to ${landedFlushes}${ofCount} landed`;
}
