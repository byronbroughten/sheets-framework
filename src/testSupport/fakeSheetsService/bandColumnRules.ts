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

export interface BandRuleShift extends BandShift {
  rule: ConditionalFormatRule;
}

export interface BandColumnRuleShift extends BandRuleShift {
  // The rule's one range, held whole across the band's rows.
  range: GridRange;
  // The columns of each Table the band holds, before the change.
  tableColumnSpans: [number, number][];
}

// Beside the band left of the change, moved with its columns, across the change point over a whole Table, or out of the band.
type RulePlacement = "before" | "moved" | "spanning" | "outside";

interface PlacedRule extends BandColumnRuleShift {
  placement: RulePlacement;
}

export const ruleOwner = "conditional-format rule";

const cellReferences = new RegExp(a1Notation.cellReference, "g");
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
  // Measured live (sheets-framework#353, Scratch353F and G): a rule the band doesn't hold keeps its range.
  shiftOutside(ruleShift: BandRuleShift): ConditionalFormatRule {
    const ranges = ruleShift.rule.ranges ?? [];
    // Several ranges give a relative reference no one origin, so one reaching the band throws, as over no bounds.
    let range: GridRange = {};
    if (ranges.length === 1) range = ranges[0] ?? {};
    return withFormulasShifted({
      ...ruleShift,
      range,
      tableColumnSpans: [],
      placement: "outside",
    });
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
  return a1Notation.mapOutsideLiterals(formula, (part) => {
    const unmeasured = unmeasuredReference.exec(part);
    if (unmeasured !== null) {
      throw unreplayedReference(unmeasured[0], placed);
    }
    return part.replace(
      cellReferences,
      (reference, anchor: string, letters: string, row: string) =>
        shiftedReference({ reference, anchor, letters, row }, placed),
    );
  });
}

interface CellReference {
  reference: string;
  anchor: string;
  letters: string;
  row: string;
}

function shiftedReference(
  cellReference: CellReference,
  placed: PlacedRule,
): string {
  const { reference, anchor, letters, row } = cellReference;
  const { change, shift, placement } = placed;
  if (placement === "before" || placement === "outside") {
    return outsideReference(cellReference, placed);
  }
  const colIndex = a1Notation.columnIndex(letters);
  if (colIndex < change.startIndex) return reference;
  const isAnchoredInSpanning = anchor === "$" && placement === "spanning";
  if (isAnchoredInSpanning) return reference;
  const movedIndex = shift.index(colIndex);
  // Measured (Scratch353C): a relative reference to a column deleted under a spanning rule stays as written too.
  const isDeletedInSpanning =
    placement === "spanning" && movedIndex === undefined;
  if (isDeletedInSpanning) return reference;
  if (anchor === "$" || placement !== "moved") {
    throw unreplayedReference(reference, placed);
  }
  if (movedIndex === undefined) return reference;
  return `${a1Notation.columnLetters(movedIndex)}${row}`;
}

// Measured (Scratch353F and G): a reference into the band follows its column, whatever its anchors, unless deleted.
function outsideReference(
  { reference, anchor, letters, row }: CellReference,
  placed: PlacedRule,
): string {
  const { change, shift, range } = placed;
  const colIndex = a1Notation.columnIndex(letters);
  const [firstCol, endCol] = referencedSpan(colIndex, anchor === "$", [
    range.startColumnIndex,
    range.endColumnIndex,
  ]);
  const rowIndex = Number(row.replace("$", "")) - 1;
  const [firstRow, endRow] = referencedSpan(rowIndex, row.startsWith("$"), [
    range.startRowIndex,
    range.endRowIndex,
  ]);
  const isLeftOfChange = endCol <= change.startIndex;
  const isOutsideBandRows =
    endRow <= change.crossStartIndex || change.crossEndIndex <= firstRow;
  if (isLeftOfChange || isOutsideBandRows) return reference;
  const isMeasuredShape =
    endCol - firstCol === 1 &&
    change.crossStartIndex <= firstRow &&
    endRow <= change.crossEndIndex &&
    (range.endColumnIndex ?? Number.MAX_SAFE_INTEGER) <= change.startIndex;
  if (!isMeasuredShape) throw unreplayedReference(reference, placed);
  const movedIndex = shift.index(colIndex);
  if (movedIndex === undefined) return reference;
  return `${anchor}${a1Notation.columnLetters(movedIndex)}${row}`;
}

// The indexes a reference reads across the rule's range: one if anchored, else one per row or column.
function referencedSpan(
  index: number,
  isAnchored: boolean,
  [start, end]: [number | undefined, number | undefined],
): [number, number] {
  if (isAnchored) return [index, index + 1];
  if (start === undefined || end === undefined) {
    return [index, Number.MAX_SAFE_INTEGER];
  }
  return [index, index + end - start];
}

function unreplayedReference(reference: string, { range }: PlacedRule): Error {
  return unreplayed(
    `a column change under the reference ${reference} in a ${ruleOwner} over ${a1Notation.range(range)}`,
  );
}
