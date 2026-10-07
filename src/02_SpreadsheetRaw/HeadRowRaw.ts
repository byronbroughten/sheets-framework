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
  addEditWarning(declaration: EditWarningDeclaration = {}): this {
    this.table.addEditWarningAt(
      this.table.rowGridRange(this.rowIndex),
      declaration,
    );
    return this;
  }
  addEditLock(declaration: EditLockDeclaration = {}): this {
    this.table.addEditLockAt(
      this.table.rowGridRange(this.rowIndex),
      declaration,
    );
    return this;
  }
  removeEditProtections(): this {
    this.table.removeEditProtectionsAt(this.table.rowGridRange(this.rowIndex));
    return this;
  }
  removeEditProtection(protection: EditProtection): this {
    this.table.removeEditProtection(protection);
    return this;
  }
}
