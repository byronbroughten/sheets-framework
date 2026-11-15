import { a1Notation } from "./a1Notation";
import type { BandChange, DimensionShift } from "./fakeGrid";
import { unreplayed } from "./unreplayed";

type ConditionalFormatRule =
  GoogleAppsScript.Sheets.Schema.ConditionalFormatRule;
type GridRange = GoogleAppsScript.Sheets.Schema.GridRange;

// A band change and where it moves each index.
export interface BandShift {
  change: BandChange;
  kind: "insert" | "remove";
  shift: DimensionShift;
}

export interface BandColumnRuleShift extends BandShift {
  rule: ConditionalFormatRule;
  // The rule's one range, held whole across the band's rows.
  range: GridRange;
  // The columns of each Table the band holds, before the change.
  tableColumnSpans: [number, number][];
}

// Left of the change, moved with its columns, or across the change point over a whole Table.
type RulePlacement = "before" | "moved" | "spanning";

interface PlacedRule extends BandColumnRuleShift {
  placement: RulePlacement;
}

export const ruleOwner = "conditional-format rule";

const cellReferences = new RegExp(a1Notation.cellReference, "g");
const stringLiteral = /("(?:[^"]|"")*")/;
// A whole-column (D:D) or whole-row (5:5) range, or a reference qualified by a sheet name.
const unmeasuredReference =
  /(?<![A-Za-z0-9_])\$?[A-Z]{1,3}:\$?[A-Z]{1,3}(?![0-9A-Za-z_(])|(?<![A-Za-z0-9_])\$?\d+:\$?\d+|!/;

// Measured live (sheets-framework#217, Scratch217A and E to G); any other shape throws rather than guess.
export const bandColumnRules = {
  shift(ruleShift: BandColumnRuleShift): ConditionalFormatRule[] {
    const { range, change, shift } = ruleShift;
    const { startColumnIndex, endColumnIndex } = range;
    if (startColumnIndex === undefined || endColumnIndex === undefined) {
      throw unreplayed(
        `a column change under a ${ruleOwner} with no column bound`,
      );
    }
    if (endColumnIndex <= change.startIndex) {
      return [withFormulasShifted({ ...ruleShift, placement: "before" })];
    }
    if (startColumnIndex < change.startIndex) {
      return [spanningRule(ruleShift, endColumnIndex)];
    }
    const movedStart = shift.index(startColumnIndex);
    const movedLast = shift.index(endColumnIndex - 1);
    if (movedStart === undefined && movedLast === undefined) return [];
    // Measured (Scratch217A): a delete of a Table's first column keeps a whole-Table rule's range too.
    if (movedStart === undefined && isOverWholeTable(ruleShift)) {
      return [spanningRule(ruleShift, endColumnIndex)];
    }
    if (movedStart === undefined || movedLast === undefined) {
      throw unreplayed(
        `a column delete over part of a ${ruleOwner} (${a1Notation.range(range)})`,
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
  endColumnIndex: number,
): ConditionalFormatRule {
  const { range, change, kind } = ruleShift;
  if (!isOverWholeTable(ruleShift)) {
    throw unreplayed(
      `a column change inside a ${ruleOwner} over part of a Table (${a1Notation.range(range)})`,
    );
  }
  const spanning = withFormulasShifted({ ...ruleShift, placement: "spanning" });
  if (kind === "remove") return spanning;
  return {
    ...spanning,
    ranges: [{ ...range, endColumnIndex: endColumnIndex + change.count }],
  };
}

function isOverWholeTable({
  range,
  tableColumnSpans,
}: BandColumnRuleShift): boolean {
  return tableColumnSpans.some(
    ([start, end]) =>
      start === range.startColumnIndex && end === range.endColumnIndex,
  );
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
  return formula
    .split(stringLiteral)
    .map((part, index) => {
      if (index % 2 === 1) return part;
      const unmeasured = unmeasuredReference.exec(part);
      if (unmeasured !== null) {
        throw unreplayedReference(unmeasured[0], placed);
      }
      return part.replace(
        cellReferences,
        (reference, anchor: string, letters: string, row: string) =>
          shiftedReference({ reference, anchor, letters, row }, placed),
      );
    })
    .join("");
}

interface CellReference {
  reference: string;
  anchor: string;
  letters: string;
  row: string;
}

function shiftedReference(
  { reference, anchor, letters, row }: CellReference,
  placed: PlacedRule,
): string {
  const { change, shift, placement } = placed;
  const colIndex = a1Notation.columnIndex(letters);
  if (colIndex < change.startIndex) return reference;
  const isAnchoredInSpanning = anchor === "$" && placement === "spanning";
  if (isAnchoredInSpanning) return reference;
  if (anchor === "$" || placement !== "moved") {
    throw unreplayedReference(reference, placed);
  }
  const movedIndex = shift.index(colIndex);
  if (movedIndex === undefined) return reference;
  return `${a1Notation.columnLetters(movedIndex)}${row}`;
}

function unreplayedReference(reference: string, { range }: PlacedRule): Error {
  return unreplayed(
    `a column change under the reference ${reference} in a ${ruleOwner} over ${a1Notation.range(range)}`,
  );
}
