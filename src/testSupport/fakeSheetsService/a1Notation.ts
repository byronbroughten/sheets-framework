type GridRange = GoogleAppsScript.Sheets.Schema.GridRange;

const stringLiteral = /("(?:[^"]|"")*")/;

export const a1Notation = {
  // A cell reference, never a function name such as LOG10(; groups: column anchor, letters, row.
  cellReference: /(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?\d+)(?![A-Za-z0-9_(])/,
  // A string literal's text is no reference, so only the parts between literals are rewritten.
  mapOutsideLiterals(
    formula: string,
    rewrite: (part: string) => string,
  ): string {
    return formula
      .split(stringLiteral)
      .map((part, index) => (index % 2 === 1 ? part : rewrite(part)))
      .join("");
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
  range(range: GridRange): string {
    const startColumn = a1Notation.columnLetters(range.startColumnIndex ?? 0);
    const lastColumn = a1Notation.columnLetters(
      (range.endColumnIndex ?? 1) - 1,
    );
    return `${startColumn}${(range.startRowIndex ?? 0) + 1}:${lastColumn}${range.endRowIndex ?? ""}`;
  },
};
