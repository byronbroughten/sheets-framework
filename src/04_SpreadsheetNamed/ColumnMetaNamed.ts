import type {
  UniformRowName,
  UniformRowValueName,
} from "../00_Source/CellValues/cellValues";
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
import type { TableName } from "../01_SpreadsheetSchema/tableConfigsTypes";
import type { ColumnMetaRaw } from "../02_SpreadsheetRaw/ColumnMetaRaw";
import type { CellIdentified } from "../03_SpreadsheetIdentified/CellIdentified";
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
    return this.table.raw.column(this.identified.colIndex);
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
  get activeColumnType(): string | undefined {
    return this.identified.activeColumnType;
  }
  updateColumnType(columnType: TableColumnType): this {
    this.identified.updateColumnType(columnType);
    return this;
  }
  uniformCell<UN extends UniformRowName>(
    rowName: UN,
  ): CellIdentified<UniformRowValueName<UN>> {
    // intentionally not cell named, because named cells only work for data...
    return this.identified.uniformCell(rowName);
  }
  prepFetchUniformCell<UN extends UniformRowName>(
    rowName: UN,
  ): CellIdentified<UniformRowValueName<UN>> {
    return this.uniformCell(rowName).prepFetch();
  }
  actionRowToDefault(): ColumnMetaNamed<TN, CN> {
    this.uniformCell("action").updateValue(false);
    return this;
  }
  addEditWarningOn(
    rowName: UniformRowName,
    declaration: EditWarningDeclaration = {},
  ): this {
    this.uniformCell(rowName).addEditWarning(declaration);
    return this;
  }
  addEditLockOn(
    rowName: UniformRowName,
    declaration: EditLockDeclaration = {},
  ): this {
    this.uniformCell(rowName).addEditLock(declaration);
    return this;
  }
}
