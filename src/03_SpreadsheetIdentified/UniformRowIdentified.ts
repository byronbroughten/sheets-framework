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
import { RowCommonIdentified } from "./ClassBases/RowCommonIdentified";
import type { TableIdentifiedProps } from "./ClassBases/TableBaseIdentified";
import { TableIdentified } from "./TableIdentified";

export type UniformRowIdentifiedProps<UN extends UniformRowName> =
  TableIdentifiedProps & { uniformRowName: UN };

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
  get sheet(): TableIdentified {
    return new TableIdentified(this.tableIdentifiedProps);
  }
  get table(): TableIdentified {
    return new TableIdentified(this.tableIdentifiedProps);
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
