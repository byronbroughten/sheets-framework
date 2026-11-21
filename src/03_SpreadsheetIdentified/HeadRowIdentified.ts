import type { CellValue } from "../00_Source/CellValues/cellValues.ts";
import type {
  EditLockDeclaration,
  EditProtection,
  EditWarningDeclaration,
} from "../00_Source/RawSource/EditProtection.ts";
import {
  type HeadRole,
  type HeadRowRoles,
  headRows,
  type HeadRowValue,
  type HeadRowValueName,
} from "../01_SpreadsheetSchema/headRows.ts";
import { HeadRowRaw } from "../02_SpreadsheetRaw/HeadRowRaw.ts";
import { RowCommonIdentified } from "./ClassBases/RowCommonIdentified.ts";
import type { TableIdentifiedProps } from "./ClassBases/TableBaseIdentified.ts";
import { TableIdentified } from "./TableIdentified.ts";

export type HeadRowIdentifiedProps<HR extends HeadRole> =
  TableIdentifiedProps & { headRole: HR };

export class HeadRowIdentified<
  HR extends HeadRole = HeadRole,
> extends RowCommonIdentified {
  declare readonly headRole: HR;
  declare readonly roles: HeadRowRoles<HR>[];
  constructor({ headRole, ...rest }: HeadRowIdentifiedProps<HR>) {
    super({ ...rest, rowIndex: headRows.index(headRole) });
    this.headRole = headRole;
    this.roles = headRows.rolesSharing(headRole);
  }
  get table(): TableIdentified {
    return new TableIdentified(this.tableIdentifiedProps);
  }
  get raw(): HeadRowRaw<HR> {
    return new HeadRowRaw({
      ...this.tableIdentifiedProps,
      headRole: this.headRole,
    });
  }
  get workingValueArr(): (HeadRowValue<HR> | "")[] {
    return this.raw.workingValueArr;
  }
  valueOrEmpty(columnId: string): CellValue<HeadRowValueName<HR>> | "" {
    return this.raw.valueOrEmpty(this.table.column(columnId).colIndex);
  }
  updateValue(columnId: string, value: HeadRowValue<HR>): this {
    this.raw.updateValue(this.table.column(columnId).colIndex, value);
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
