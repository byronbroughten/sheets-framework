import { Val } from "@byronbroughten/utils/val";

export interface FormulaText {
  tableKey: string;
  formula: string;
}

export interface Brace {
  text: string;
  tableKey: string;
  columnName: string;
  // `@` before the column name: this row's value rather than the whole column.
  isThisRow: boolean;
}

export interface BraceMatch {
  text: string;
  literal: string | undefined;
  firstAt: string | undefined;
  first: string | undefined;
  secondAt: string | undefined;
  second: string | undefined;
  // A `${…}` the brace pattern didn't read.
  unread: string | undefined;
}

// A sheet reference's title, quoted or bare, and the lookbehind that keeps it from starting mid-word.
const sheetTitle = /'(?:[^']|'')*'|[A-Za-z_][\w.]*/.source;
const beforeSheetPrefix = /(?<![\w$.'])/.source;

export const formulaPatterns = {
  stringLiteral: /"(?:[^"]|"")*"/.source,
  brace: /\$\{(@?)([^{}[\]"@]+?)(?:\[(@?)([^{}[\]"@]+)\])?\}/.source,
  // Any `${…}`, read as a brace or not.
  anyBrace: /\$\{[^{}"]*\}/.source,
  liveReference: /([A-Za-z_][\w.]*)\[([^[\]"]+)\]/.source,
  sheetPrefix: `${beforeSheetPrefix}(${sheetTitle})!`,
  // The lookarounds keep a function name such as LOG10( from reading as a cell.
  a1Reference: `${beforeSheetPrefix}((?:${sheetTitle})!)?${
    /(\$?)([A-Z]{1,3})(\$?)(\d+)?(?::(\$?)([A-Z]{1,3})(\$?)(\d+)?)?(?![\w$(])/
      .source
  }`,
} as const;

// A string literal is matched first and kept, so a brace inside one stays text; any other `${…}` falls to the last group.
export function bracePattern(): RegExp {
  return new RegExp(
    `(${formulaPatterns.stringLiteral})|${formulaPatterns.brace}|(${formulaPatterns.anyBrace})`,
    "g",
  );
}

// The match text and bracePattern's capture groups, counted off an empty match so a new group moves it too.
export const braceGroupCount = Val.assert(
  new RegExp(`${bracePattern().source}|`).exec(""),
  "an empty match of the brace pattern",
).length;

export function replaceBraces(
  { tableKey, formula }: FormulaText,
  replace: (brace: Brace) => string,
): string {
  return formula.replace(bracePattern(), (...groups: unknown[]) => {
    const match = braceMatchOf(groups);
    const brace = braceOf(tableKey, match);
    if (brace === undefined) return match.text;
    return replace(brace);
  });
}

// Takes bracePattern's groups in order, from a replace callback's arguments or a matchAll result.
export function braceMatchOf(groups: ArrayLike<unknown>): BraceMatch {
  const [text, literal, firstAt, first, secondAt, second, unread] = Array.from(
    groups,
  ).map((group) => (typeof group === "string" ? group : undefined));
  return {
    text: text ?? "",
    literal,
    firstAt,
    first,
    secondAt,
    second,
    unread,
  };
}

// Undefined for a string literal, which isn't a brace, for an unread `${…}`, and for `@` before a Table key, which marks no column.
export function braceOf(
  tableKey: string,
  { text, literal, firstAt, first, secondAt, second, unread }: BraceMatch,
): Brace | undefined {
  if (literal !== undefined || unread !== undefined) return undefined;
  if (first === undefined) return undefined;
  const isFirstAt = firstAt === "@";
  if (second === undefined) {
    return { text, tableKey, columnName: first, isThisRow: isFirstAt };
  }
  if (isFirstAt) return undefined;
  return {
    text,
    tableKey: first,
    columnName: second,
    isThisRow: secondAt === "@",
  };
}

export function braceMessage(
  brace: Brace,
  problem: string,
  formula: string,
): string {
  return `The brace ${brace.text} ${problem}, in formula ${formula}.`;
}

export function noColumnMessage(brace: Brace, formula: string): string {
  return braceMessage(
    brace,
    `names no column of Table ${brace.tableKey}`,
    formula,
  );
}
