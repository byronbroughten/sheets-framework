import { Val } from "@byronbroughten/utils/val";

import type { HeldGroup } from "../01_SpreadsheetSchema/configReaders/dryRunHeld.ts";
import type {
  ForcedTrait,
  SyncSubject,
} from "../01_SpreadsheetSchema/configReaders/syncRunResult.ts";
import type {
  HeldCell,
  HeldCellItem,
  HeldItem,
  HeldRow,
} from "./FlushClassifier/flushClassifierTypes.ts";

type OtherItem = Exclude<HeldItem, HeldCellItem>;
type OtherItemOf<KN extends OtherItem["kind"]> = Extract<
  OtherItem,
  { kind: KN }
>;

// The report lists this many cells of a group; the held file lists every one.
const listedCellLimit = 20;

// Each flush's held items, and the ones its send snapshot plan leaves to version history.
export interface HeldFlushes {
  flushes: HeldItem[][];
  versionHistoryItems: HeldItem[][];
}

interface FlaggedCells {
  isUndoneOnlyByVersionHistory: boolean;
  items: HeldCellItem[];
}

// The dry run's held items as the report groups them: cells per Table, column, kind and version-history flag, and one group per other item.
export function heldGroups({
  flushes,
  versionHistoryItems,
}: HeldFlushes): HeldGroup[] {
  const cellGroups = flushes.reduce((groups, items, flushIndex) => {
    const flagged = new Set(versionHistoryItems[flushIndex] ?? []);
    for (const item of items.filter(isHeldCellItem)) {
      const isUndoneOnlyByVersionHistory = flagged.has(item);
      const key = JSON.stringify([
        isUndoneOnlyByVersionHistory,
        item.kind,
        item.cell.sheetGid,
        item.cell.table,
        item.cell.column,
      ]);
      const group = groups.get(key) ?? {
        isUndoneOnlyByVersionHistory,
        items: [],
      };
      group.items.push(item);
      groups.set(key, group);
    }
    return groups;
  }, new Map<string, FlaggedCells>());
  return [
    ...[...cellGroups.values()].map(cellGroup),
    ...flushes
      .flat()
      .flatMap((item): HeldGroup[] =>
        isHeldCellItem(item) ? [] : [otherGroup(item)],
      ),
  ];
}

// Each forced trait as its own group; a forced send needs no token, so these sit outside the held block.
export function forcedTraitGroups(traits: ForcedTrait[]): HeldGroup[] {
  return traits.map((trait) => otherGroup({ kind: "forcedTrait", trait }));
}

export function isHeldCellItem(item: HeldItem): item is HeldCellItem {
  return item.kind === "enteredData" || item.kind === "unreadCell";
}

const cellGroupWhat: Record<
  HeldCellItem["kind"],
  (cells: number, cellWord: string) => string
> = {
  enteredData: (cells, cellWord) =>
    `the send overwrites or deletes entered data in ${cells} ${cellWord}:`,
  unreadCell: (cells, cellWord) =>
    `the send writes ${cells} ${cellWord} no read covered, so they may hold entered data:`,
};

function cellGroup({
  isUndoneOnlyByVersionHistory,
  items,
}: FlaggedCells): HeldGroup {
  const { cell, kind } = Val.assert(items[0], "held cell group's first cell");
  const cells = items.length;
  const cellWord = cells === 1 ? "cell" : "cells";
  const listed = items.slice(0, listedCellLimit).map(cellLine);
  const unlisted = cells - listed.length;
  const lines =
    unlisted > 0
      ? [
          ...listed,
          `… and ${unlisted} more in ${cell.column}, counted, not listed`,
        ]
      : listed;
  const group: HeldGroup = {
    subject: cellSubject(cell),
    what: cellGroupWhat[kind](cells, cellWord),
    cells,
    lines,
  };
  if (!isUndoneOnlyByVersionHistory) return group;
  return { ...group, isUndoneOnlyByVersionHistory };
}

// A cell on a sheet no read covered has no Table, so its tab's GID names it.
function cellSubject({ table, column, sheetGid }: HeldCell): SyncSubject {
  if (table === undefined) return { tab: `gid ${sheetGid}`, path: column };
  return { table, column };
}

function cellLine(item: HeldCellItem): string {
  const row = rowLabel(item.cell.row);
  if (item.kind === "unreadCell") return `${row}  not read`;
  const { change } = item;
  if (change.kind === "delete") return `${row}  "${change.before}"  → deleted`;
  return `${row}  "${change.before}"  → "${change.after}"`;
}

function rowLabel(row: HeldRow): string {
  if ("rowId" in row) return row.rowId;
  const name = row.nameValue === undefined ? "" : ` (${row.nameValue})`;
  return `row ${row.sheetRow}${name}`;
}

const otherGroupByKind: {
  [KN in OtherItem["kind"]]: (item: OtherItemOf<KN>) => HeldGroup;
} = {
  dropdownReset: ({ table }) => ({
    subject: { table },
    what: 'dropdown style reset (dropdownStyleReset: "ask")',
    lines: [],
  }),
  forcedTrait: ({ trait: { table, column, trait } }) => ({
    subject: { table, column, path: trait },
    what: "sent over the live value by --force",
    lines: [],
  }),
  findReplace: ({ find, replacement }) => ({
    what: `a find-and-replace of "${find}" with "${replacement}", which can reach entered data anywhere`,
    lines: [],
  }),
  unmeasured: ({ verb, detail }) => ({
    what: `a ${verb} request the gate couldn't measure: ${detail}`,
    lines: [],
  }),
  unclassified: ({ verb }) => ({
    what: `a ${verb} request the gate can't judge`,
    lines: [],
  }),
};

function otherGroup<KN extends OtherItem["kind"]>(
  item: OtherItemOf<KN> & { kind: KN },
): HeldGroup {
  const build: (item: OtherItemOf<KN>) => HeldGroup =
    otherGroupByKind[item.kind];
  return { ...build(item), isUndoneOnlyByVersionHistory: true };
}
