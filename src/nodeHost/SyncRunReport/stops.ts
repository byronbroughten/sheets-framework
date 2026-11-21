import {
  beforeFirstColumnText,
  type SyncBeforeFirstColumnStop,
  type SyncConflict,
  type SyncHeaderTakenStop,
  type SyncHeadRowsStop,
  type SyncNoSuchRuleStop,
  type SyncStop,
  type SyncTableNameTakenStop,
  type SyncTabTitleRepeatedStop,
  type SyncTabTitleTakenStop,
  type SyncUnmodelableRuleStop,
} from "../../01_SpreadsheetSchema/configReaders/syncRunResult";
import { headerZone } from "../../01_SpreadsheetSchema/headerZone";
import { headRows } from "../../01_SpreadsheetSchema/headRows";
import { reportText } from "../reportText";

const { indent, counted, subjectLabel } = reportText;

const layout = {
  rowLabelWidth: 30,
  wayOutWidth: 6,
} as const;

export const stopLines = {
  // Why the run stopped, for the result line, in the result line's order.
  why(stops: SyncStop[]): string {
    return stopKinds.flatMap((kind) => stopReason(kind, stops)).join(" and ");
  },
  blocks(stops: SyncStop[]): string[] {
    return stops.flatMap((stop, index) => [
      ...(index === 0 ? [] : [""]),
      ...stopBlock(stop, index + 1),
    ]);
  },
  resolveLine(stops: SyncStop[], command: string): string {
    const [stop] = stops;
    if (stops.length !== 1 || stop === undefined) {
      return `Resolve each stop above, then rerun ${command}.`;
    }
    return `Resolve ${stopWords[stop.kind].name} 1, then rerun ${command}.`;
  },
};

type StopKind = SyncStop["kind"];

type StopOf<SK extends StopKind> = Extract<SyncStop, { kind: SK }>;

interface StopWords<SK extends StopKind> {
  // What the resolve line calls one.
  name: string;
  // Why the run stopped, for the result line.
  reason(stops: StopOf<SK>[]): string;
  block(stop: StopOf<SK>, number: number): string[];
}

// The result line's order.
const stopKinds: StopKind[] = [
  "conflict",
  "headRows",
  "headerTaken",
  "beforeFirstColumn",
  "tableNameTaken",
  "tabTitleTaken",
  "tabTitleRepeated",
  "noSuchRule",
  "unmodelableRule",
];

const stopWords: { [SK in StopKind]: StopWords<SK> } = {
  conflict: {
    name: "conflict",
    reason({ length }) {
      return counted(length, "conflict", "conflicts");
    },
    block: conflictBlock,
  },
  headRows: {
    name: "stop",
    // A grant stops once per Table its head rows would sit on, but counts as one Table.
    reason(stops) {
      const tables = new Set(stops.map(({ table }) => table));
      return `${counted(tables.size, "Table", "Tables")} can't take head rows`;
    },
    block: headRowsBlock,
  },
  headerTaken: {
    name: "stop",
    reason({ length }) {
      return `${counted(length, "new column", "new columns")} taken`;
    },
    block: headerTakenBlock,
  },
  beforeFirstColumn: {
    name: "stop",
    reason({ length }) {
      return `${counted(length, "new column", "new columns")} before a first column`;
    },
    block: beforeFirstColumnBlock,
  },
  tableNameTaken: {
    name: "stop",
    reason({ length }) {
      return `${counted(length, "new Table name", "new Table names")} taken`;
    },
    block: tableNameTakenBlock,
  },
  tabTitleTaken: {
    name: "stop",
    reason({ length }) {
      return `${counted(length, "new tab title", "new tab titles")} taken`;
    },
    block: tabTitleTakenBlock,
  },
  tabTitleRepeated: {
    name: "stop",
    reason({ length }) {
      return `${counted(length, "new tab title", "new tab titles")} repeated`;
    },
    block: tabTitleRepeatedBlock,
  },
  noSuchRule: {
    name: "stop",
    reason({ length }) {
      return `${counted(length, "rule", "rules")} not found`;
    },
    block: noSuchRuleBlock,
  },
  unmodelableRule: {
    name: "stop",
    reason({ length }) {
      return `${counted(length, "unmodelable rule", "unmodelable rules")} to replace`;
    },
    block: unmodelableRuleBlock,
  },
};

// The one phrase for a stop's entry left pending.
const entryStaysText = "The entry stays in alterations.ts.";

// A stop on a Table's entry is headed by the entry's path; a new Table's or tab's, which has none, by its name.
function stopHeading(number: number, table: string, path: string): string {
  return `STOP ${number} · ${table} › ${path}`;
}

function stopReason<SK extends StopKind>(
  kind: SK,
  stops: SyncStop[],
): string[] {
  const ofKind = stops.filter((stop): stop is StopOf<SK> => stop.kind === kind);
  if (ofKind.length === 0) return [];
  const words: StopWords<SK> = stopWords[kind];
  return [words.reason(ofKind)];
}

function stopBlock<SK extends StopKind>(
  stop: StopOf<SK>,
  number: number,
): string[] {
  const words: StopWords<SK> = stopWords[stop.kind];
  return words.block(stop, number);
}

function headRowsBlock(stop: SyncHeadRowsStop, number: number): string[] {
  return [
    stopHeading(number, stop.table, "letApiAccess"),
    ...indent(2, headRowsReason(stop)),
  ];
}

function headRowsReason(stop: SyncHeadRowsStop): string[] {
  const { table, tab } = stop;
  if ("neighbour" in stop) {
    return [
      `${table}'s head rows, ${headRows.aboveHeaderLabel} on its tab, ${tab}, would sit on Table "${stop.neighbour}".`,
      `Move ${table} or "${stop.neighbour}" so that ${table}'s head rows sit clear of it, then rerun.`,
      entryStaysText,
    ];
  }
  const { headerRow } = stop;
  const headerPlace = headerZone.headerRowsLabel;
  return [
    `${table}'s header row is on row ${headerRow} of its tab, ${tab}. A Table needs its header on ${headerPlace} to take head rows,`,
    `${headRows.aboveHeaderLabel}. Move ${table} so its header sits on ${headerPlace}, then rerun.`,
    entryStaysText,
  ];
}

function headerTakenBlock(
  { table, header, column, columnId }: SyncHeaderTakenStop,
  number: number,
): string[] {
  return [
    stopHeading(number, table, "newColumns"),
    ...indent(2, [
      `${table} already has a column headed "${header}": ${column}, ${columnId}.`,
      entryStaysText,
    ]),
  ];
}

function beforeFirstColumnBlock(
  stop: SyncBeforeFirstColumnStop,
  number: number,
): string[] {
  return [
    stopHeading(number, stop.table, stop.path),
    ...indent(2, [
      beforeFirstColumnText(stop),
      "Name a later column in before, or leave before out to add it at the Table's end.",
      entryStaysText,
    ]),
  ];
}

function tableNameTakenBlock(
  { name, existing, tableId }: SyncTableNameTakenStop,
  number: number,
): string[] {
  return [
    `STOP ${number} · new Table "${name}"`,
    ...indent(2, [
      `A Table already gives its name's codebase name: "${existing}", ${tableId}.`,
      entryStaysText,
    ]),
  ];
}

function tabTitleTakenBlock(
  { title, existing, gid }: SyncTabTitleTakenStop,
  number: number,
): string[] {
  return [
    `STOP ${number} · new tab "${title}"`,
    ...indent(2, [
      `A tab already gives its title's codebase name: "${existing}", GID ${gid}.`,
      entryStaysText,
    ]),
  ];
}

function tabTitleRepeatedBlock(
  { title, paths }: SyncTabTitleRepeatedStop,
  number: number,
): string[] {
  return [
    `STOP ${number} · new tab "${title}"`,
    ...indent(2, [
      `${paths.join(", ")} give one title's codebase name. Give each new tab its own title, or merge the entries.`,
      "The entries stay in alterations.ts.",
    ]),
  ];
}

function noSuchRuleBlock(
  { table, path, handles }: SyncNoSuchRuleStop,
  number: number,
): string[] {
  return [
    stopHeading(number, table, path),
    ...indent(2, [
      "no such rule: adopt by removing the entry.",
      `No rule on ${table} has this handle now: a person changed the rule after the last sync`,
      "(a handle comes from its columns, formula, condition and format).",
      ...ruleHandleLines(table, handles),
      entryStaysText,
    ]),
  ];
}

function ruleHandleLines(table: string, handles: string[]): string[] {
  if (handles.length === 0) return [`${table} has no rules now.`];
  return [`${table}'s rules now, by handle:`, ...indent(2, handles)];
}

function unmodelableRuleBlock(
  { table, path, handle }: SyncUnmodelableRuleStop,
  number: number,
): string[] {
  return [
    stopHeading(number, table, path),
    ...indent(2, [
      `The rule ${handle} is unmodelable: sync can't write a rule like it, so it can only be removed.`,
      "Move the handle to conditionalFormats.remove and add the rule you want instead.",
      entryStaysText,
    ]),
  ];
}

function conflictBlock(conflict: SyncConflict, number: number): string[] {
  const { subject, label, cause, yours, agreed, live } = conflict;
  const rows = [
    ["alterations.ts (yours)", yours],
    ["configs (last agreed)", agreed],
    ["live sheet", live],
  ].map(([name = "", value = ""]) => name.padEnd(layout.rowLabelWidth) + value);
  return [
    `CONFLICT ${number} · ${subjectLabel(subject)} · ${label}`,
    ...indent(2, [...rows, cause, "Ways out:"]),
    ...indent(
      4,
      waysOut(conflict).map(
        ([way, text]) => `${way.padEnd(layout.wayOutWidth)} ${text}`,
      ),
    ),
  ];
}

const forceNeeds =
  "Needs the developer's yes for this trait; the yes covers one run.";

function waysOut(conflict: SyncConflict): [string, string][] {
  const { yours, live } = conflict;
  // A deleted column leaves no live value to set the entry to.
  const adopt: [string, string] = [
    "adopt",
    conflict.isColumnGone === true
      ? "Remove the entry. Free, but your change is given up."
      : `Remove the entry, or set it to ${live}. Free, but your change to this trait is given up.`,
  ];
  const ask: [string, string] = [
    "ask",
    "Change nothing and ask the developer which value they want.",
  ];
  const forceFlag = forceFlagOf(conflict);
  if (forceFlag === undefined) return [adopt, ask];
  return [
    adopt,
    [
      "force",
      `Send ${yours} over the sheet's ${live}. ${forceNeeds} ${forceFlag}`,
    ],
    ask,
  ];
}

// --force addresses <table>/<column>/<trait>, so a Table's or tab's own trait, or a deleted column's, has no force to offer.
function forceFlagOf({
  subject,
  isColumnGone,
}: SyncConflict): string | undefined {
  if (isColumnGone === true) return undefined;
  if (!("column" in subject)) return undefined;
  const { table, column, path } = subject;
  if (column === undefined || path === undefined) return undefined;
  return `Flag: --force ${table}/${column}/${path}`;
}
