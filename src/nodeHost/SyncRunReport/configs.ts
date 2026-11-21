import type {
  SyncConfigs,
  SyncMovedName,
} from "../../01_SpreadsheetSchema/configReaders/syncRunResult";
import { reportText } from "../reportText";

const { indent, subjectLabel } = reportText;

export const configsLines = {
  // What the refresh changed, in the Configs section's order.
  changes(configs: SyncConfigs, generatedDir: string): string[] {
    return [
      ...heldBackLines(configs),
      ...renamedLines(configs),
      ...namesMovedLines(configs, generatedDir),
      ...rewriteLines(configs),
      ...typedTraitLines(configs),
      ...configs.notes,
    ];
  },
};

function heldBackLines({ heldBack }: SyncConfigs): string[] {
  if (heldBack.length === 0) return [];
  return [
    "Held back while a conflict is open, so the next run still sees it:",
    ...indent(
      2,
      heldBack.map(
        ({ subject, value }) =>
          `${subjectLabel(subject)}: the configs keep ${value} for this trait until the conflict is resolved.`,
      ),
    ),
  ];
}

function renamedLines({ renamed }: SyncConfigs): string[] {
  if (renamed.length === 0) return [];
  return [
    "Renamed by a person since the last sync; the codebase name follows the header:",
    ...indent(
      2,
      renamed.map(
        ({ table, from, to, headerFrom, headerTo }) =>
          `${table} › ${from} → ${to}    header "${headerFrom}" → "${headerTo}"`,
      ),
    ),
  ];
}

function namesMovedLines(
  { namesMoved }: SyncConfigs,
  generatedDir: string,
): string[] {
  if (namesMoved.length === 0) return [];
  return [
    `Names that now point at a different ID than at the last sync. Reread ${generatedDir}idNameMap.ts before writing another entry:`,
    ...indent(
      2,
      namesMoved.map(
        (moved) =>
          `${movedNameLabel(moved)}    was ${moved.wasId}, now ${moved.nowId}`,
      ),
    ),
  ];
}

const movedNameKinds: Record<SyncMovedName["kind"], string> = {
  table: "Table",
  tab: "tab",
  column: "column",
};

function movedNameLabel(moved: SyncMovedName): string {
  const kind = movedNameKinds[moved.kind];
  if (moved.kind === "column") {
    return `${kind} ${moved.table} › ${moved.column}`;
  }
  return `${kind} ${moved.name}`;
}

function rewriteLines({ rewrites }: SyncConfigs): string[] {
  if (rewrites.length === 0) return [];
  return [
    "Rewritten in alterations.ts, so each entry still names the column it named before:",
    ...indent(
      2,
      rewrites.map(({ from, to, why }) => {
        const reason = why === undefined ? "" : `    (${why})`;
        return `${from} → ${to}${reason}`;
      }),
    ),
  ];
}

function typedTraitLines({ typedTraitChanges }: SyncConfigs): string[] {
  if (typedTraitChanges.length === 0) return [];
  return [
    "Typed traits a person changed since the last sync, on Tables with API access. tsc fails where code relied on the old ones:",
    ...indent(
      2,
      typedTraitChanges.map(
        ({ table, column, trait, from, to }) =>
          `${table} › ${column}    ${trait} ${from} → ${to}`,
      ),
    ),
  ];
}
