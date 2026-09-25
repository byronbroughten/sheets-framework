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
import { uniformRows } from "../01_SpreadsheetSchema/uniformRows";
import { UniformRowRaw } from "../02_SpreadsheetRaw/UniformRowRaw";
import type { StrictOmit } from "../utils/Obj";
import type { RowIdentifiedProps } from "./ClassBases/RowBaseIdentified";
import { RowCommonIdentified } from "./ClassBases/RowCommonIdentified";
import { SheetIdentified } from "./SheetIdentified";

export interface UniformRowIdentifiedProps<
  UN extends UniformRowName,
> extends StrictOmit<RowIdentifiedProps, "rowIndex"> {
  uniformRowName: UN;
}

export class UniformRowIdentified<
  UN extends UniformRowName = UniformRowName,
> extends RowCommonIdentified {
  readonly uniformRowName: UN;
  constructor({ uniformRowName, ...rest }: UniformRowIdentifiedProps<UN>) {
    super({
      ...rest,
      rowIndex: uniformRows.index(uniformRowName),
    });
    this.uniformRowName = uniformRowName;
    this.schema.validateUniformRowIndex(this.rowIndex, this.uniformRowName);
  }
  get sheet(): SheetIdentified {
    return new SheetIdentified(this.sheetIdentifiedProps);
  }
  get raw(): UniformRowRaw<UN> {
    return new UniformRowRaw({
      ...this.rowIdentifiedProps,
      uniformRowName: this.uniformRowName,
    });
  }
  get valueName(): UniformRowValueName<UN> {
    return this.schema.uniformValueName(this.uniformRowName);
  }
  valueOrEmpty(columnId: string): UniformRowValue<UN> | "" {
    return this.raw.valueOrEmpty(this.sheet.column(columnId).colIndex);
  }
  get activeValueArr(): (UniformRowValue<UN> | "")[] {
    return this.raw.activeValueArr;
  }
  updateValue(columnId: string, value: UniformRowValue<UN>): this {
    this.raw.updateValue(this.sheet.column(columnId).colIndex, value);
    return this;
  }
  addEditWarning(declaration: EditWarningDeclaration = {}): this {
    this.raw.addEditWarning(declaration);
    return this;
  }
  addEditLock(declaration: EditLockDeclaration = {}): this {
    this.raw.addEditLock(declaration);
    return this;
  }
  removeEditProtections(): this {
    this.raw.removeEditProtections();
    return this;
  }
  removeEditProtection(protection: EditProtection): this {
    this.raw.removeEditProtection(protection);
    return this;
  }
}
