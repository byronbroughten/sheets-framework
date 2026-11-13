import type { SyncHeldGroup } from "../01_SpreadsheetSchema/configReaders/syncRunResult";
import type {
  HeldCell,
  HeldItem,
  HeldRow,
} from "./FlushClassifier/flushClassifierTypes";

type CellItem = Extract<HeldItem, { kind: "enteredData" | "unreadCell" }>;

// The report lists this many cells of a group; the held file lists every one.
const listedCellLimit = 20;

// The dry run's held items as the report groups them: cells per Table, column and kind, and one group per other item.
export function heldGroups(flushes: HeldItem[][]): SyncHeldGroup[] {
  const items = flushes.flat();
  const cellGroups = items
    .filter(isCellItem)
    .reduce<Map<string, CellItem[]>>((groups, item) => {
      const key = JSON.stringify([
        item.kind,
        item.cell.sheetGid,
        item.cell.table,
        item.cell.column,
      ]);
      groups.set(key, [...(groups.get(key) ?? []), item]);
      return groups;
    }, new Map());
  return [
    ...[...cellGroups.values()].map(cellGroup),
    ...items.flatMap((item): SyncHeldGroup[] =>
      isCellItem(item) ? [] : [otherGroup(item)],
    ),
  ];
}

function isCellItem(item: HeldItem): item is CellItem {
  return item.kind === "enteredData" || item.kind === "unreadCell";
}

function cellGroup(items: CellItem[]): SyncHeldGroup {
  const [first] = items;
  if (first === undefined) throw new Error("A held cell group has no cells.");
  const { cell, kind } = first;
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
  if (kind === "enteredData") {
    return {
      ...cellSubject(cell),
      what: `the send overwrites or deletes entered data in ${cells} ${cellWord}:`,
      cells,
      lines,
    };
  }
  return {
    ...cellSubject(cell),
    what: `the send writes ${cells} ${cellWord} no read covered, so they may hold entered data:`,
    cells,
    lines,
  };
}

// A cell on a sheet no read covered has no Table, so its tab's GID names it.
function cellSubject({
  table,
  column,
  sheetGid,
}: HeldCell): Pick<SyncHeldGroup, "subject"> {
  if (table === undefined) {
    return { subject: { tab: `gid ${sheetGid}`, path: column } };
  }
  return { subject: { table, column } };
}

function cellLine(item: CellItem): string {
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

function otherGroup(item: Exclude<HeldItem, CellItem>): SyncHeldGroup {
  if (item.kind === "dropdownReset") {
    return {
      subject: { table: item.table },
      what: 'dropdown style reset (dropdownStyleReset: "ask")',
      lines: [],
    };
  }
  if (item.kind === "forcedTrait") {
    const { table, column, trait } = item.trait;
    return {
      subject: { table, column, path: trait },
      what: "sent over the live value by --force",
      lines: [],
    };
  }
  if (item.kind === "findReplace") {
    return {
      what: `a find-and-replace of "${item.find}" with "${item.replacement}", which can reach entered data anywhere`,
      lines: [],
    };
  }
  if (item.kind === "unmeasured") {
    return {
      what: `a ${item.verb} request the gate couldn't measure: ${item.detail}`,
      lines: [],
    };
  }
  return {
    what: `a ${item.verb} request the gate can't judge`,
    lines: [],
  };
}
