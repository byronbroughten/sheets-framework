import type { CellValue } from "../../00_Source/CellValues/cellValues";
import { RowBaseIdentified } from "./RowBaseIdentified";

export abstract class RowCommonIdentified extends RowBaseIdentified {
  abstract get workingValueArr(): CellValue[];
  hasValue(value: unknown): boolean {
    return this.workingValueArr.includes(value as CellValue);
  }
  prepFetchFull(): void {
    this.fetchTargets.push({
      kind: "fullRow",
      row: this.rowIndex,
    });
  }
}
