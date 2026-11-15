import { a1Columns } from "./a1Columns";
import type { BandChange, DimensionShift } from "./fakeGrid";

type ConditionalFormatRule =
  GoogleAppsScript.Sheets.Schema.ConditionalFormatRule;
type GridRange = GoogleAppsScript.Sheets.Schema.GridRange;

export interface BandColumnRuleShift {
  rule: ConditionalFormatRule;
  // The rule's one range, held whole across the band's rows.
  range: GridRange;
  change: BandChange;
  kind: "insert" | "remove";
  shift: DimensionShift;
  // The columns of each Table the band holds, before the change.
  tableColumnSpans: [number, number][];
}

// Left of the change, moved with its columns, or across the change point over a whole Table.
type RulePlacement = "before" | "moved" | "spanning";

interface PlacedRule extends BandColumnRuleShift {
  placement: RulePlacement;
}

// A cell reference outside a string literal, never a function name such as LOG10(.
const cellReference =
  /(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?\d+)(?![A-Za-z0-9_(])/g;
const stringLiteral = /("(?:[^"]|"")*")/;

// Measured live (sheets-framework#217, Scratch217E to G); any other shape throws rather than guess.
export const bandColumnRules = {
  shift(ruleShift: BandColumnRuleShift): ConditionalFormatRule[] {
    const { range, change, shift } = ruleShift;
    const { startColumnIndex, endColumnIndex } = range;
    if (startColumnIndex === undefined || endColumnIndex === undefined) {
      throw unreplayed(
        "a column change under a conditional-format rule with no column bound",
      );
    }
    if (endColumnIndex <= change.startIndex) {
      return [withFormulasShifted({ ...ruleShift, placement: "before" })];
    }
    if (startColumnIndex < change.startIndex) {
      return [spanningRule(ruleShift, startColumnIndex, endColumnIndex)];
    }
    const movedStart = shift.index(startColumnIndex);
    const movedLast = shift.index(endColumnIndex - 1);
    if (movedStart === undefined && movedLast === undefined) return [];
    if (movedStart === undefined || movedLast === undefined) {
      throw unreplayed(
        `a column delete over part of a conditional-format rule (${rangeA1(range)})`,
      );
    }
    const moved = withFormulasShifted({ ...ruleShift, placement: "moved" });
    return [
      {
        ...moved,
        ranges: [
          {
            ...range,
            startColumnIndex: movedStart,
            endColumnIndex: movedLast + 1,
          },
        ],
      },
    ];
  },
};

// Measured: an insert grows a rule over a whole Table, and a delete leaves it one column past the Table.
function spanningRule(
  ruleShift: BandColumnRuleShift,
  startColumnIndex: number,
  endColumnIndex: number,
): ConditionalFormatRule {
  const { range, change, kind, tableColumnSpans } = ruleShift;
  const isOverWholeTable = tableColumnSpans.some(
    ([start, end]) => start === startColumnIndex && end === endColumnIndex,
  );
  if (!isOverWholeTable) {
    throw unreplayed(
      `a column change inside a conditional-format rule over part of a Table (${rangeA1(range)})`,
    );
  }
  const spanning = withFormulasShifted({ ...ruleShift, placement: "spanning" });
  if (kind === "remove") return spanning;
  return {
    ...spanning,
    ranges: [{ ...range, endColumnIndex: endColumnIndex + change.count }],
  };
}

function withFormulasShifted(placed: PlacedRule): ConditionalFormatRule {
  const { rule } = placed;
  const condition = rule.booleanRule?.condition;
  if (condition?.values === undefined) return rule;
  const values = condition.values.map((value) => {
    const formula = value.userEnteredValue;
    if (formula?.startsWith("=") !== true) return value;
    return { ...value, userEnteredValue: shiftedFormula(formula, placed) };
  });
  return {
    ...rule,
    booleanRule: { ...rule.booleanRule, condition: { ...condition, values } },
  };
}

// Measured: an $-anchored column stays as written; a relative one moves with its column unless deleted.
function shiftedFormula(formula: string, placed: PlacedRule): string {
  const { change, shift, placement, range } = placed;
  return formula
    .split(stringLiteral)
    .map((part, index) => {
      if (index % 2 === 1) return part;
      return part.replace(
        cellReference,
        (reference, anchor: string, letters: string, row: string) => {
          const colIndex = a1Columns.index(letters);
          if (colIndex < change.startIndex) return reference;
          const isAnchoredInSpanning =
            anchor === "$" && placement === "spanning";
          if (isAnchoredInSpanning) return reference;
          if (anchor === "$" || placement !== "moved") {
            throw unreplayed(
              `a column change under the reference ${reference} in a conditional-format rule over ${rangeA1(range)}`,
            );
          }
          const movedIndex = shift.index(colIndex);
          if (movedIndex === undefined) return reference;
          return `${a1Columns.letters(movedIndex)}${row}`;
        },
      );
    })
    .join("");
}

function rangeA1(range: GridRange): string {
  const startColumn = a1Columns.letters(range.startColumnIndex ?? 0);
  const lastColumn = a1Columns.letters((range.endColumnIndex ?? 1) - 1);
  return `${startColumn}${(range.startRowIndex ?? 0) + 1}:${lastColumn}${range.endRowIndex ?? ""}`;
}

function unreplayed(what: string): Error {
  return new Error(`The fake Sheets service does not replay ${what}.`);
}
