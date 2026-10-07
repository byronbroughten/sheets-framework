import {
  type CodebaseNameDelimiter,
  codebaseNameDelimiter,
  getUniformRowValueName,
  type UniformRowName,
  type UniformRowValueName,
} from "../00_Source/CellValues/cellValues";
import type {
  SheetColIndex,
  SheetRowIndex,
} from "../00_Source/RawSource/SheetIndex";
import { Obj } from "../utils/Obj";
import { Str } from "../utils/Str";
import { tableLayout } from "./tableLayout";
import { uniformRows } from "./uniformRows";

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
  uniformValueName<UN extends UniformRowName>(
    name: UN,
  ): UniformRowValueName<UN> {
    return getUniformRowValueName(name);
  }
  get uniformRowNames(): UniformRowName[] {
    return Obj.keys(uniformRows.indexes());
  }
  uniformRowIndex(name: UniformRowName): number {
    return uniformRows.index(name);
  }
  uniformRowNameByIndex(rowIndex: number): UniformRowName {
    const uniformRowName = uniformRows.nameByIndex().get(rowIndex);
    if (!uniformRowName) {
      throw new Error(
        `Row index ${rowIndex} does not correspond to a known uniform row name.`,
      );
    }
    return uniformRowName;
  }
  isUniformRowIndex(rowIndex: number, rowName?: UniformRowName): boolean {
    const isUniform = uniformRows.nameByIndex().has(rowIndex);
    if (rowName) {
      return isUniform && this.uniformRowNameByIndex(rowIndex) === rowName;
    } else {
      return isUniform;
    }
  }
  validateUniformRowIndex(rowIndex: number, rowName?: UniformRowName): void {
    if (!this.isUniformRowIndex(rowIndex, rowName)) {
      throw new Error(
        `Row index ${rowIndex} is not a uniform row. Uniform rows are: ${Obj.keys(
          uniformRows.indexes(),
        )
          .map((name) => `${name} (index ${uniformRows.indexes()[name]})`)
          .join(", ")}`,
      );
    }
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
    return uniformRows.indexes().columnId;
  }
  get tableHeaderRowIndex(): number {
    return uniformRows.indexes().header;
  }
  get actionRowIndex(): number {
    return uniformRows.indexes().action;
  }
  get idDelimiter(): string {
    return tableLayout.idDelimiter;
  }
}
