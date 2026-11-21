import {
  braceGroupCount,
  type BraceMatch,
  braceMatchOf,
  braceOf,
  bracePattern,
  formulaPatterns,
  type FormulaText,
} from "./braces";
import type { ConfiguredColumn } from "./cellFormula";

export interface NamedColumn extends ConfiguredColumn {
  // The brace as written, so a refusal can quote it.
  brace: string;
}

export interface RenamedTable {
  name: string;
  // Each column's name now, by its previous one.
  columns: Map<string, string>;
}

// Each Table by its previous name: `names` as braces write them, `liveNames` as Sheets' `Table[Header]` does.
export interface FormulaRenames {
  names: Map<string, RenamedTable>;
  liveNames: Map<string, RenamedTable>;
}

interface FormulaRenameProps {
  // Undefined for a Table this run creates, whose own columns no regen has renamed.
  tableKey: string | undefined;
  formula: string;
  renames: FormulaRenames;
}

interface BraceRename {
  match: BraceMatch;
  ownTable: RenamedTable | undefined;
  names: Map<string, RenamedTable>;
}

interface LiveReferenceMatch {
  liveTable: string;
  header: string;
}

interface LiveReferenceRename extends LiveReferenceMatch {
  liveNames: Map<string, RenamedTable>;
}

export const formulaNames = {
  columnsNamed({ tableKey, formula }: FormulaText): NamedColumn[] {
    return [...formula.matchAll(bracePattern())].flatMap((groups) => {
      const brace = braceOf(tableKey, braceMatchOf(groups));
      if (brace === undefined) return [];
      return [
        {
          brace: brace.text,
          tableKey: brace.tableKey,
          columnName: brace.columnName,
        },
      ];
    });
  },
  // The Sheets name of each Table a `Table[Header]` reference outside a string literal or brace names.
  liveTablesNamed({ formula }: Pick<FormulaText, "formula">): string[] {
    const pattern = new RegExp(
      `${bracePattern().source}|${formulaPatterns.liveReference}`,
      "g",
    );
    return [...formula.matchAll(pattern)].flatMap((groups) => {
      const reference = liveReferenceMatchOf([...groups]);
      if (reference === undefined) return [];
      return [reference.liveTable];
    });
  },
  // The title each sheet reference outside a string literal or brace names.
  sheetsNamed({ formula }: Pick<FormulaText, "formula">): string[] {
    const { stringLiteral, anyBrace, sheetPrefix } = formulaPatterns;
    const pattern = new RegExp(
      `(${stringLiteral}|${anyBrace})|${sheetPrefix}`,
      "g",
    );
    return [...formula.matchAll(pattern)].flatMap(([, skipped, title]) => {
      if (skipped !== undefined || title === undefined) return [];
      if (!title.startsWith("'")) return [title];
      return [title.slice(1, -1).replaceAll("''", "'")];
    });
  },
  // Each brace and `Table[Header]` reference outside a string literal, under the names a regen gave it.
  renamed({ tableKey, formula, renames }: FormulaRenameProps): string {
    const ownTable =
      tableKey === undefined ? undefined : renames.names.get(tableKey);
    const pattern = new RegExp(
      `${bracePattern().source}|${formulaPatterns.liveReference}`,
      "g",
    );
    return formula.replace(pattern, (...groups: unknown[]) => {
      const reference = liveReferenceMatchOf(groups);
      if (reference !== undefined) {
        return renamedLiveReference({
          ...reference,
          liveNames: renames.liveNames,
        });
      }
      return renamedBrace({
        match: braceMatchOf(groups),
        ownTable,
        names: renames.names,
      });
    });
  },
  // Each `${…}` outside a string literal that reads as no brace, such as `${@table[column]}`, so a run can refuse it rather than paste it.
  malformedBraces({ formula }: Pick<FormulaText, "formula">): string[] {
    return [...formula.matchAll(bracePattern())].flatMap((groups) => {
      const match = braceMatchOf(groups);
      if (match.literal !== undefined) return [];
      if (braceOf("", match) !== undefined) return [];
      return [match.text];
    });
  },
};

// Keeps the brace's form as written, bare or qualified, `@` included.
function renamedBrace({ match, ownTable, names }: BraceRename): string {
  const { text, literal, firstAt, first, secondAt, second, unread } = match;
  if (literal !== undefined || unread !== undefined) return text;
  if (first === undefined) return text;
  if (second === undefined) {
    const column = ownTable?.columns.get(first) ?? first;
    return `\${${firstAt ?? ""}${column}}`;
  }
  if (firstAt === "@") return text;
  const table = names.get(first);
  if (table === undefined) return text;
  const column = table.columns.get(second) ?? second;
  return `\${${table.name}[${secondAt ?? ""}${column}]}`;
}

// Takes the groups a live reference adds after bracePattern's; undefined when the match was the brace pattern's.
function liveReferenceMatchOf(
  groups: unknown[],
): LiveReferenceMatch | undefined {
  const [liveTable, header] = groups.slice(braceGroupCount);
  if (typeof liveTable !== "string" || typeof header !== "string") {
    return undefined;
  }
  return { liveTable, header };
}

function renamedLiveReference({
  liveNames,
  liveTable,
  header,
}: LiveReferenceRename): string {
  const table = liveNames.get(liveTable);
  if (table === undefined) return `${liveTable}[${header}]`;
  return `${table.name}[${table.columns.get(header) ?? header}]`;
}
