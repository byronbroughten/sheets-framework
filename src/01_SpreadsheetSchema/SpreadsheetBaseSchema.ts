import { Str } from "@byronbroughten/utils/str";

import {
  type CodebaseNameDelimiter,
  codebaseNameDelimiter,
} from "../00_Source/CellValues/cellValues";
import type {
  SheetColIndex,
  SheetRowIndex,
} from "../00_Source/RawSource/SheetIndex";
import { headRows } from "./headRows";
import { tableLayout } from "./tableLayout";


export class SpreadsheetBaseSchema {
  get codebaseNameDelimiter(): CodebaseNameDelimiter {
    return codebaseNameDelimiter;
  }
  combineNames<SA extends string, SB extends string>(
    name1: SA,
    name2: SB,
  ): `${SA}${CodebaseNameDelimiter}${SB}` {
    return `${name1}${this.codebaseNameDelimiter}${name2}`;
  }
  get idHeader(): string {
    return tableLayout.idHeader;
  }
  get nameHeader(): string {
    return tableLayout.nameHeader;
  }
  titleToName(sheetTitle: string): string {
    return Str.sentenceToCamelCase(sheetTitle);
  }
  positionLabel(rowIndex: SheetRowIndex, colIndex: SheetColIndex): string {
    return `row ${rowIndex + 1}, column ${this.columnLetter(colIndex)}`;
  }
  columnLetter(colIndex: SheetColIndex): string {
    let letters = "";
    let remaining: number = colIndex;
    while (remaining >= 0) {
      letters = String.fromCharCode(65 + (remaining % 26)) + letters;
      remaining = Math.floor(remaining / 26) - 1;
    }
    return letters;
  }
  anchoredA1(colIndex: SheetColIndex, rowIndex: SheetRowIndex): string {
    return `$${this.columnLetter(colIndex)}${rowIndex + 1}`;
  }
  get colIdRowIndex(): number {
    return headRows.index("columnId");
  }
  get tableHeaderRowIndex(): number {
    return headRows.index("header");
  }
  get actionRowIndex(): number {
    return headRows.index("action");
  }
  get idDelimiter(): string {
    return tableLayout.idDelimiter;
  }
}
