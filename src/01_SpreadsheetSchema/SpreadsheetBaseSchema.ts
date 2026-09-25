import {
  type CodebaseNameDelimiter,
  codebaseNameDelimiter,
  getUniformRowValueName,
  type UniformRowName,
  type UniformRowValueName,
} from "../00_Source/CellValues/cellValues";
import { Obj } from "../utils/Obj";
import { Str } from "../utils/Str";
import {
  type LiveSpreadsheetConfig,
  ssConfigGet,
} from "./spreadsheetConfigTypes";
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
  get idHeader(): LiveSpreadsheetConfig["idHeader"] {
    return ssConfigGet("idHeader");
  }
  get nameHeader(): LiveSpreadsheetConfig["nameHeader"] {
    return ssConfigGet("nameHeader");
  }
  titleToName(sheetTitle: string): string {
    return Str.sentenceToCamelCase(sheetTitle);
  }
  uniformValueName<UN extends UniformRowName>(
    name: UN,
  ): UniformRowValueName<UN> {
    return getUniformRowValueName(name);
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
  isTableStart(startRowIndex: number, startColumnIndex: number): boolean {
    return (
      startRowIndex === this.tableHeaderRowIndex &&
      startColumnIndex === this.startTableColIndex
    );
  }
  validateTableStart(startRowIndex: number, startColumnIndex: number): void {
    if (!this.isTableStart(startRowIndex, startColumnIndex)) {
      throw new Error(
        `A Table starting at ${this.positionLabel(
          startRowIndex,
          startColumnIndex,
        )} must start at ${this.tableStartLabel}.`,
      );
    }
  }
  get tableStartLabel(): string {
    return this.positionLabel(
      this.tableHeaderRowIndex,
      this.startTableColIndex,
    );
  }
  positionLabel(rowIndex: number, colIndex: number): string {
    return `row ${rowIndex + 1}, column ${this.columnLetter(colIndex)}`;
  }
  columnLetter(colIndex: number): string {
    let letters = "";
    let remaining = colIndex;
    while (remaining >= 0) {
      letters = String.fromCharCode(65 + (remaining % 26)) + letters;
      remaining = Math.floor(remaining / 26) - 1;
    }
    return letters;
  }
  anchoredA1(colIndex: number, rowIndex: number): string {
    return `$${this.columnLetter(colIndex)}${rowIndex + 1}`;
  }
  isDataRowIndex(rowIndex: number): boolean {
    return rowIndex >= this.topDataRowIdx;
  }
  get startTableColIndex(): number {
    return ssConfigGet("startTableColIndexBase0");
  }
  get colIdRowIndex(): number {
    return uniformRows.indexes().columnId;
  }
  get tableHeaderRowIndex(): number {
    return uniformRows.indexes().tableHeader;
  }
  get actionRowIndex(): number {
    return uniformRows.indexes().action;
  }
  get topDataRowIdx(): number {
    return this.tableHeaderRowIndex + 1;
  }
  get idDelimiter(): string {
    return ssConfigGet("idDelimiter");
  }
}
