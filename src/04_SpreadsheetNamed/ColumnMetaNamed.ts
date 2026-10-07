import type {
  ColumnFullName,
  ColumnName,
  ColumnValueName,
  MakeColumnFullName,
} from "../01_SpreadsheetSchema/columnConfigsTypes";
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
}
