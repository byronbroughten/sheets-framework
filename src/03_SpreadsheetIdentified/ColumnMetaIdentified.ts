import type { TableColumnType } from "../00_Source/RawSource/RawSource";
import type { ValueName, VnToCvn } from "../01_SpreadsheetSchema/valueSchemas";
import { ColumnMetaRaw } from "../02_SpreadsheetRaw/ColumnMetaRaw";
import { ColumnCommonIdentified } from "./ClassBases/ColumnCommonIdentified";
import { ColumnIdentified } from "./ColumnIdentified";
import { SheetMetaIdentified } from "./SheetMetaIdentified";

export class ColumnMetaIdentified<
  VN extends ValueName = ValueName,
> extends ColumnCommonIdentified<VN> {
  get raw(): ColumnMetaRaw<VnToCvn<VN>> {
    return new ColumnMetaRaw({
      ...this.tableIdentifiedProps,
      colIndex: this.colIndex,
    });
  }
  get table(): SheetMetaIdentified {
    return new SheetMetaIdentified(this.tableIdentifiedProps);
  }
  get primary(): ColumnIdentified<VN> {
    return new ColumnIdentified(this.columnIdentifiedProps);
  }
  updateColumnType(columnType: TableColumnType): this {
    this.raw.updateColumnType(columnType);
    return this;
  }
}
