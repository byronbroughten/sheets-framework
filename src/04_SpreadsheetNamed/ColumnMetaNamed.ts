import type {
  EditLockDeclaration,
  EditWarningDeclaration,
} from "../00_Source/RawSource/EditProtection";
import type { TableColumnType } from "../00_Source/RawSource/RawSource";
import type {
  ColumnFullName,
  ColumnName,
  ColumnValueName,
  MakeColumnFullName,
} from "../01_SpreadsheetSchema/columnConfigsTypes";
import type { HeadRole } from "../01_SpreadsheetSchema/headRows";
import type { TableName } from "../01_SpreadsheetSchema/tableConfigsTypes";
import type { ColumnMetaRaw } from "../02_SpreadsheetRaw/ColumnMetaRaw";
import { ColumnMetaIdentified } from "../03_SpreadsheetIdentified/ColumnMetaIdentified";
import { ColumnCommonNamed } from "./ClassBases/ColumnCommonNamed";
import { ColumnNamed } from "./ColumnNamed";
import { SheetMetaNamed } from "./SheetMetaNamed";

export class ColumnMetaNamed<
  TN extends TableName,
  CN extends ColumnName<TN> = ColumnName<TN>,
> extends ColumnCommonNamed<TN, CN> {
  get table(): SheetMetaNamed<TN> {
    return new SheetMetaNamed(this.sheetNamedProps);
  }
  get raw(): ColumnMetaRaw {
    return this.table.raw.primary.column(this.identified.colIndex).meta;
  }
  get identified(): ColumnMetaIdentified<ColumnValueName<TN, CN>> {
    return new ColumnMetaIdentified<ColumnValueName<TN, CN>>({
      ...this.table.identified.tableIdentifiedProps,
      columnId: this.columnId,
    });
  }
  get primary(): ColumnNamed<TN, CN> {
    return new ColumnNamed(this.columnNamedProps);
  }
  get colIndex(): number {
    return this.identified.colIndex;
  }
  get fullName(): MakeColumnFullName<TN, CN> & ColumnFullName {
    return this.schema.fullName;
  }
  updateColumnType(columnType: TableColumnType): this {
    this.identified.updateColumnType(columnType);
    return this;
  }
  actionRowToDefault(): ColumnMetaNamed<TN, CN> {
    this.primary.headCell("action").updateValue(false);
    return this;
  }
  addEditWarningOn(
    headRole: HeadRole,
    declaration: EditWarningDeclaration = {},
  ): this {
    this.primary.headCell(headRole).addEditWarning(declaration);
    return this;
  }
  addEditLockOn(
    headRole: HeadRole,
    declaration: EditLockDeclaration = {},
  ): this {
    this.primary.headCell(headRole).addEditLock(declaration);
    return this;
  }
}
