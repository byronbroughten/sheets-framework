import type {
  UniformRowName,
  UniformRowValue,
  UniformRowValueName,
} from "../00_Source/CellValues/cellValues";
import type {
  EditLockDeclaration,
  EditProtection,
  EditWarningDeclaration,
} from "../00_Source/RawSource/EditProtection";
import { UniformRowBaseRaw } from "./ClassBases/UniformRowBaseRaw";

export class UniformRowRaw<
  UN extends UniformRowName = UniformRowName,
> extends UniformRowBaseRaw<UN> {
  valueOrEmpty(colIndex: number): UniformRowValue<UN> | "" {
    return this.cell<UniformRowValueName<UN>>(colIndex).valueOrEmpty();
  }
  updateValue(colIndex: number, value: UniformRowValue<UN>): this {
    this.cell(colIndex).updateValue(value);
    return this;
  }
  addEditWarning(declaration: EditWarningDeclaration = {}): this {
    this.sheet.addEditWarningAt(
      this.sheet.rowGridRange(this.rowIndex),
      declaration,
    );
    return this;
  }
  addEditLock(declaration: EditLockDeclaration = {}): this {
    this.sheet.addEditLockAt(
      this.sheet.rowGridRange(this.rowIndex),
      declaration,
    );
    return this;
  }
  removeEditProtections(): this {
    this.sheet.removeEditProtectionsAt(this.sheet.rowGridRange(this.rowIndex));
    return this;
  }
  removeEditProtection(protection: EditProtection): this {
    this.sheet.removeEditProtection(protection);
    return this;
  }
}
