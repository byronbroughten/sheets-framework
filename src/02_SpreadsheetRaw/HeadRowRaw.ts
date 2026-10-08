import type { CellValue } from "../00_Source/CellValues/cellValues";
import type {
  EditLockDeclaration,
  EditProtection,
  EditWarningDeclaration,
} from "../00_Source/RawSource/EditProtection";
import type {
  HeadRole,
  HeadRowValue,
  HeadRowValueName,
} from "../01_SpreadsheetSchema/headRows";
import { HeadRowBaseRaw } from "./ClassBases/HeadRowBaseRaw";
import type { TableGridRange } from "./ClassTypes/StateRaw";

export class HeadRowRaw<
  HR extends HeadRole = HeadRole,
> extends HeadRowBaseRaw<HR> {
  valueOrEmpty(colIndex: number): CellValue<HeadRowValueName<HR>> | "" {
    return this.cell<HeadRowValueName<HR>>(colIndex).valueOrEmpty();
  }
  updateValue(colIndex: number, value: HeadRowValue<HR>): this {
    this.cell(colIndex).updateValue(value);
    return this;
  }
  // The whole sheet row, not just the Table's columns.
  get tableGridRange(): TableGridRange {
    return { startRowIndex: this.rowIndex, endRowIndex: this.rowIndex + 1 };
  }
  addEditWarning(declaration: EditWarningDeclaration = {}): this {
    this.table.addEditWarningAt(this.tableGridRange, declaration);
    return this;
  }
  addEditLock(declaration: EditLockDeclaration = {}): this {
    this.table.addEditLockAt(this.tableGridRange, declaration);
    return this;
  }
  removeEditProtections(): this {
    this.table.removeEditProtectionsAt(this.tableGridRange);
    return this;
  }
  removeEditProtection(protection: EditProtection): this {
    this.table.sheet.removeEditProtection(protection);
    return this;
  }
}
