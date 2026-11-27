import type {
  DryRunHeld,
  HeldBlock,
  HeldGroup,
} from "../../01_SpreadsheetSchema/configReaders/dryRunHeld.ts";
import type { HeldItem } from "../FlushClassifier/flushClassifierTypes.ts";
import { heldGroups } from "../heldGroups.ts";
import { ReportText } from "../ReportText.ts";
import { gateSectionTitles } from "./gateSectionTitles.ts";

const { section, indent, formatCount, counted, subjectLabel } = ReportText;

export const heldReport = {
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
