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
    if (/^[A-Za-z_]\w*$/.test(title)) return title;
    return `'${title.replaceAll("'", "''")}'`;
  },
  range(range: GridRange): string {
    const startColumn = a1Notation.columnLetters(range.startColumnIndex ?? 0);
    const lastColumn = a1Notation.columnLetters(
      (range.endColumnIndex ?? 1) - 1,
    );
    return `${startColumn}${(range.startRowIndex ?? 0) + 1}:${lastColumn}${range.endRowIndex ?? ""}`;
  },
};
