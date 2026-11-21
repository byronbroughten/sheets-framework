import { sendableTraits } from "../../01_SpreadsheetSchema/configReaders/SyncEntryRefuser";
import type {
  SyncClashingPlacements,
  SyncEntry,
  SyncRefusal,
} from "../../01_SpreadsheetSchema/configReaders/syncRunResult";
import { ReportText } from "../ReportText";

const { indent, subjectLabel } = ReportText;

export const refusalLines = {
  blocks(refusals: SyncRefusal[]): string[] {
    return refusals.flatMap((refusal, index) => {
      const { heading, lines } = refusalWordsOf(refusal);
      return [`REFUSED ${index + 1} · ${heading}`, ...indent(2, lines)];
    });
  },
};

interface RefusalWords {
  heading: string;
  lines: string[];
}

type RefusalKind = SyncRefusal["kind"];

type RefusalOf<RK extends RefusalKind> = Extract<SyncRefusal, { kind: RK }>;

function refusalWordsOf<RK extends RefusalKind>(
  refusal: RefusalOf<RK>,
): RefusalWords {
  const words: (refusal: RefusalOf<RK>) => RefusalWords =
    refusalWords[refusal.kind];
  return words(refusal);
}

// A refusal on a Table's entry is headed by the entry's path.
function tableRefusal(
  { table, path }: { table: string; path: string },
  lines: string[],
): RefusalWords {
  return { heading: `${table} › ${path}`, lines };
}

const refusalWords: {
  [RK in RefusalKind]: (refusal: RefusalOf<RK>) => RefusalWords;
} = {
  touchesDeletion({ deleted, deletedAt, touchedBy, takenTable }) {
    return {
      heading: subjectLabel(deleted),
      lines: [
        `${deletedAt} deletes ${subjectLabel(deleted)}, and ${touchedBy} touches it in the same run.`,
        ...(takenTable === undefined
          ? []
          : [`The tab's deletion already takes the Table ${takenTable}.`]),
        "Take one of the two out, then rerun.",
      ],
    };
  },
  noSuchTable({ entry, name }) {
    const noun = entryNouns[entry];
    return {
      heading: `${entry} › ${name}`,
      lines: [
        `no such ${noun}: adopt by removing the entry.`,
        `${noMapEntryText(noun, name)}, so it is gone from the sheet or went by another name at the last sync.`,
      ],
    };
  },
  notSendableYet({ entry, name, path }) {
    return {
      heading: `${entryNouns[entry]} ${name} › ${path}`,
      lines: [
        `sync can't send ${path} yet: for now it sends only a Table's ${[...sendableTraits.table].join(", ")}, a column's ${[...sendableTraits.column].join(", ")} and a tab's ${[...sendableTraits.tab].join(", ")}.`,
        "Take this entry out, then rerun.",
      ],
    };
  },
  noSuchBeforeTab(refusal) {
    return tableRefusal(refusal, [
      `${refusal.path} names ${refusal.before}, which is neither a tab at the last sync nor a newTabs title.`,
      "Name a tab the ID-name map holds or a newTabs title, then rerun.",
    ]);
  },
  noApiAccess(refusal) {
    const { table } = refusal;
    return tableRefusal(refusal, [
      `${table} has no API access, so sync has nothing to compare it against: grant \`letApiAccess\` first.`,
      ...grantRunLines(
        table,
        `take this entry out, put letApiAccess: true on ${table}`,
      ),
    ]);
  },
  besideGrant(refusal) {
    const { table } = refusal;
    return tableRefusal(refusal, [
      `${table} has no API access yet, so sync has nothing to compare its columns against.`,
      ...grantRunLines(
        table,
        `keep letApiAccess: true on ${table}, take this entry out`,
      ),
    ]);
  },
  malformedBrace(refusal) {
    return tableRefusal(refusal, [
      `${refusal.brace} isn't a brace: write \${@column} or \${table[@column]} for this row, \${column} or \${table[column]} for the whole column.`,
    ]);
  },
  unknownColumn(refusal) {
    return tableRefusal(refusal, [
      `${refusal.column} names no column of ${refusal.table}, nor one its newColumns create. Its columns: ${refusal.columns.join(", ")}.`,
    ]);
  },
  repeatedNewColumn(refusal) {
    return tableRefusal(refusal, [
      `${refusal.path} gives the column name ${refusal.column}, as ${refusal.firstPath} does: the two columns can't both be named it.`,
      newColumnNameOwnText,
    ]);
  },
  removedAndReplaced(refusal) {
    return tableRefusal(refusal, [
      `${refusal.removedAt} removes the rule this handle names, and ${refusal.path} replaces it in the same run.`,
      keepOneText,
    ]);
  },
  clashingPlacements(refusal) {
    return tableRefusal(refusal, [
      clashingPlacementsText[refusal.clash](refusal),
      keepOneText,
    ]);
  },
  renamedColumn(refusal) {
    return tableRefusal(refusal, [
      `${refusal.column} is the name this run's rename gives ${refusal.oldName}, and alterations.ts names a column by its old name until the run lands.`,
      `Use the column's old name, ${refusal.oldName}, this run, then rerun.`,
    ]);
  },
  unknownBraceTable(refusal) {
    return tableRefusal(refusal, [
      `${refusal.brace} names no Table: ${noMapEntryText("Table", refusal.braceTable)}.`,
    ]);
  },
  unknownBrace(refusal) {
    return tableRefusal(refusal, [
      `${refusal.brace} names no column of ${refusal.braceTable}. Its columns: ${refusal.columns.join(", ")}.`,
    ]);
  },
};

const entryNouns: Record<SyncEntry, string> = {
  tables: "Table",
  tabs: "tab",
  newTabs: "new tab",
};

// Why the two paths of a `moves` or `newColumns` clash can't both hold.
const clashingPlacementsText: Record<
  SyncClashingPlacements["clash"],
  (refusal: SyncClashingPlacements) => string
> = {
  sameBefore: ({ path, otherPath, column }) =>
    `${path} puts its group before ${column}, as ${otherPath} does: only one group can sit right before it.`,
  repeatedColumn: ({ path, otherPath, column }) =>
    `${path} moves ${column}, as ${otherPath} does: a column can sit in only one group.`,
  movedBefore: ({ path, otherPath, column }) =>
    `${path} puts its group before ${column}, which ${otherPath} moves: a group's target can't itself move.`,
  sameNewColumnsBefore: ({ path, otherPath, column }) =>
    `${path} puts new columns before ${column}, as ${otherPath} does: write them as one group, in the order they go.`,
  newColumnsBeforeMoveTarget: ({ path, otherPath, column }) =>
    `${path} puts new columns before ${column}, as ${otherPath} puts its group: only one can sit right before it.`,
  newColumnsBeforeMovedColumn: ({ path, otherPath, column }) =>
    `${path} puts new columns before ${column}, which ${otherPath} moves: new columns' before can't itself move.`,
};

// The one phrase for two entries that can't both land.
const keepOneText = "Keep one of the two, then rerun.";

// The one phrase for two new headers giving one column name.
const newColumnNameOwnText =
  "Give each a header whose column name is its own, then rerun.";

// The one phrase for a name the previous ID-name map lacks.
function noMapEntryText(noun: string, name: string): string {
  return `idNameMap.ts holds no ${noun} named ${name}`;
}

function grantRunLines(table: string, steps: string): string[] {
  return [
    `Granting access takes its own run: ${steps}, and run sync.`,
    "Other Tables' entries may share that run.",
    `Put this entry back in a later run, once ${table}'s columns are in idNameMap.ts.`,
  ];
}
