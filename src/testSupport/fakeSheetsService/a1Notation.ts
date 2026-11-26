type GridRange = GoogleAppsScript.Sheets.Schema.GridRange;

export interface CellReference {
  reference: string;
  anchor: string;
  letters: string;
  row: string;
}

const stringLiteral = /("(?:[^"]|"")*")/;
// A cell reference, never a function name such as LOG10(; groups: column anchor, letters, row.
const cellReference =
  /(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?\d+)(?![A-Za-z0-9_(])/;
const cellReferences = new RegExp(cellReference, "g");

const plainSheetName = /[A-Za-z_]\w*/.source;
const wholePlainSheetName = new RegExp(`^${plainSheetName}$`);
const address = /([A-Z]{1,3})([1-9]\d*)?/.source;
// Groups: quoted title, plain title, start column, start row, end column, end row.
const sheetRange = new RegExp(
  `^(?:'((?:[^']|'')+)'|(${plainSheetName}))!${address}(?::${address})?$`,
);

export interface SheetRangeParts {
  title: string;
  startColumn: string;
  startRow: string | undefined;
  endColumn: string | undefined;
  endRow: string | undefined;
}

export const a1Notation = {
  cellReference,
  outsideLiterals(formula: string): string[] {
    return formula.split(stringLiteral).filter((_, index) => index % 2 === 0);
  },
  // A string literal's text is no reference, so only the parts between literals are rewritten.
  mapCellReferences(
    formula: string,
    rewrite: (cellReference: CellReference) => string,
  ): string {
    return formula
      .split(stringLiteral)
      .map((part, index) => {
        if (index % 2 === 1) return part;
        return part.replace(
          cellReferences,
          (reference, anchor: string, letters: string, row: string) =>
            rewrite({ reference, anchor, letters, row }),
        );
      })
      .join("");
  },
  cellReferencesIn(text: string): CellReference[] {
    return [...text.matchAll(cellReferences)].map(
      ([reference, anchor = "", letters = "A", row = ""]) => ({
        reference,
        anchor,
        letters,
        row,
      }),
    );
  },
  columnLetters(colIndex: number): string {
    const letter = String.fromCharCode(65 + (colIndex % 26));
    if (colIndex < 26) return letter;
    return a1Notation.columnLetters(Math.floor(colIndex / 26) - 1) + letter;
  },
  columnIndex(letters: string): number {
    return (
      [...letters].reduce(
        (index, letter) => index * 26 + letter.charCodeAt(0) - 64,
        0,
      ) - 1
    );
  },
  // Quoting a title that isn't a plain identifier was never probed live.
  sheetName(title: string): string {
    if (wholePlainSheetName.test(title)) return title;
    return `'${title.replaceAll("'", "''")}'`;
  },
  // A range naming its sheet and a column; undefined for any other shape.
  sheetRangeParts(text: string): SheetRangeParts | undefined {
    const match = sheetRange.exec(text);
    if (match === null) return undefined;
    const [, quoted, plain = "", startColumn = "", startRow] = match;
    const [, , , , , endColumn, endRow] = match;
    // A lone column with no row, such as Title!A, is no A1 range.
    if (endColumn === undefined && startRow === undefined) return undefined;
    const title = quoted?.replaceAll("''", "'") ?? plain;
    return { title, startColumn, startRow, endColumn, endRow };
  },
  // As Google echoes a range: after its sheet's name, a single cell named once.
  qualifiedRange(title: string, range: GridRange): string {
    const cells = a1Notation.range(range);
    const [first, last] = cells.split(":");
    const echoed = first === last ? first : cells;
    return `${a1Notation.sheetName(title)}!${echoed}`;
  },
  range(range: GridRange): string {
    const startColumn = a1Notation.columnLetters(range.startColumnIndex ?? 0);
    const lastColumn = a1Notation.columnLetters(
      (range.endColumnIndex ?? 1) - 1,
    );
    return `${startColumn}${(range.startRowIndex ?? 0) + 1}:${lastColumn}${range.endRowIndex ?? ""}`;
  },
};
