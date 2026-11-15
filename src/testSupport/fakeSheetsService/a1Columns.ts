export const a1Columns = {
  letters(colIndex: number): string {
    const letter = String.fromCharCode(65 + (colIndex % 26));
    if (colIndex < 26) return letter;
    return a1Columns.letters(Math.floor(colIndex / 26) - 1) + letter;
  },
  index(letters: string): number {
    return (
      [...letters].reduce(
        (index, letter) => index * 26 + letter.charCodeAt(0) - 64,
        0,
      ) - 1
    );
  },
};
