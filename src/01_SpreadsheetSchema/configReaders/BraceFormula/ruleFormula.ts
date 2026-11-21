import {
  type SheetColIndex,
  SheetIndex,
  type SheetRowIndex,
} from "../../../00_Source/RawSource/SheetIndex";
import { SpreadsheetBaseSchema } from "../../SpreadsheetBaseSchema";
import { braces, formulaPatterns, type FormulaText } from "./braces";

const schema = new SpreadsheetBaseSchema();

export interface RuleLayout {
  headerTopLeft: { rowIndex: SheetRowIndex; colIndex: SheetColIndex };
  // The Table's exclusive end, so the last body row is the row just above it.
  endRowIndex: SheetRowIndex;
  // Undefined at a column with no ID yet, which a run that mints none leaves out of the configs.
  columnNamesLeftToRight: readonly (string | undefined)[];
}

interface RuleFormulaProps extends FormulaText {
  layout: RuleLayout;
}
// No Table key: a rule reaches only its own Table, so every brace it gives back is bare.
export interface RuleDecompileProps extends Omit<RuleFormulaProps, "tableKey"> {
  isOneColumn: boolean;
}

interface RuleReference {
  sheetPrefix: string | undefined;
  start: CellAddress;
  end: CellAddress | undefined;
}

interface CellAddress {
  isColAbsolute: boolean;
  colLetter: string;
  isRowAbsolute: boolean;
  rowNumber: number | undefined;
}

interface RuleBraceContext {
  columnNamesByLetter: Map<string, string>;
  firstBodyRowNumber: number;
  isOneColumn: boolean;
}

export const ruleFormula = {
  // Takes a layout rather than reading the configs, so a rule can compile against a planned layout.
  compileForRule({ tableKey, formula, layout }: RuleFormulaProps): string {
    return braces.replace({ tableKey, formula }, (brace) => {
      if (brace.tableKey !== tableKey) {
        throw new Error(
          braces.message(
            brace,
            "names another Table, which a rule's formula can't reach",
            formula,
          ),
        );
      }
      const offset = layout.columnNamesLeftToRight.indexOf(brace.columnName);
      if (offset === -1) {
        throw new Error(braces.noColumnMessage(brace, formula));
      }
      const colIndex = SheetIndex.col(layout.headerTopLeft.colIndex + offset);
      const firstBodyRow = SheetIndex.row(layout.headerTopLeft.rowIndex + 1);
      if (brace.isThisRow) return schema.anchoredA1(colIndex, firstBodyRow);
      return schema.anchoredColumnA1({
        colIndex,
        firstRowIndex: firstBodyRow,
        lastRowIndex: SheetIndex.row(layout.endRowIndex - 1),
      });
    });
  },
  // The inverse of compileForRule; undefined when a reference has no brace form.
  decompileForRule({
    formula,
    layout,
    isOneColumn,
  }: RuleDecompileProps): string | undefined {
    const columnNamesByLetter = layout.columnNamesLeftToRight.reduce(
      (byLetter, columnName, offset) => {
        if (columnName === undefined) return byLetter;
        const colIndex = SheetIndex.col(layout.headerTopLeft.colIndex + offset);
        byLetter.set(schema.columnLetter(colIndex), columnName);
        return byLetter;
      },
      new Map<string, string>(),
    );
    const firstBodyRowNumber = layout.headerTopLeft.rowIndex + 2;
    const pattern = new RegExp(
      `(${formulaPatterns.stringLiteral})|${formulaPatterns.a1Reference}`,
      "g",
    );
    let isExpressible = true;
    const decompiled = formula.replace(
      pattern,
      (
        match: string,
        literal: string | undefined,
        sheetPrefix: string | undefined,
        ...addressParts: unknown[]
      ) => {
        if (literal !== undefined) return match;
        const reference = ruleReferenceOf(sheetPrefix, addressParts);
        if (reference === undefined) return match;
        const brace = ruleBraceOf(reference, {
          columnNamesByLetter,
          firstBodyRowNumber,
          isOneColumn,
        });
        if (brace === undefined) isExpressible = false;
        return brace ?? match;
      },
    );
    if (!isExpressible) return undefined;
    return decompiled;
  },
};

// Undefined for a bare word such as TRUE, which has neither a row nor a colon.
function ruleReferenceOf(
  sheetPrefix: string | undefined,
  addressParts: unknown[],
): RuleReference | undefined {
  const start = cellAddressOf(addressParts.slice(0, 4));
  if (start === undefined) return undefined;
  const end = cellAddressOf(addressParts.slice(4, 8));
  if (end === undefined && start.rowNumber === undefined) return undefined;
  return { sheetPrefix, start, end };
}

// Takes one address's four groups: column $, column letters, row $, row digits.
function cellAddressOf(groups: unknown[]): CellAddress | undefined {
  const [colAbs, colLetter, rowAbs, rowDigits] = groups;
  if (typeof colLetter !== "string") return undefined;
  let rowNumber: number | undefined;
  if (typeof rowDigits === "string") rowNumber = Number(rowDigits);
  return {
    isColAbsolute: colAbs === "$",
    colLetter,
    isRowAbsolute: rowAbs === "$",
    rowNumber,
  };
}

function ruleBraceOf(
  { sheetPrefix, start, end }: RuleReference,
  context: RuleBraceContext,
): string | undefined {
  if (sheetPrefix !== undefined) return undefined;
  const columnName = context.columnNamesByLetter.get(start.colLetter);
  if (columnName === undefined) return undefined;
  if (end === undefined) {
    if (start.isRowAbsolute) return undefined;
    if (start.rowNumber !== context.firstBodyRowNumber) return undefined;
    if (!start.isColAbsolute && !context.isOneColumn) return undefined;
    return `\${@${columnName}}`;
  }
  const isBodyColumn =
    start.isColAbsolute &&
    start.isRowAbsolute &&
    start.rowNumber === context.firstBodyRowNumber &&
    end.isColAbsolute &&
    end.colLetter === start.colLetter &&
    isBodyColumnEnd(end, context.firstBodyRowNumber);
  if (!isBodyColumn) return undefined;
  return `\${${columnName}}`;
}

// Any anchored end row from the first body row down, so a Table that grew since, or a rule from before ends were bounded, keeps its brace.
function isBodyColumnEnd(
  { isRowAbsolute, rowNumber }: CellAddress,
  firstBodyRowNumber: number,
): boolean {
  if (rowNumber === undefined) return true;
  return isRowAbsolute && rowNumber >= firstBodyRowNumber;
}
