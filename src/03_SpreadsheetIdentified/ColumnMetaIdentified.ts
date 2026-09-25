import type {
  UniformRowName,
  UniformRowValueName,
} from "../00_Source/CellValues/cellValues";
import type { TableColumnType } from "../00_Source/RawSource/RawSource";
import type { ValueName, VnToCvn } from "../01_SpreadsheetSchema/valueSchemas";
import { ColumnMetaRaw } from "../02_SpreadsheetRaw/ColumnMetaRaw";
import { CellIdentified } from "./CellIdentified";
import { ColumnCommonIdentified } from "./ClassBases/ColumnCommonIdentified";
import { ColumnIdentified } from "./ColumnIdentified";
import { SheetMetaIdentified } from "./SheetMetaIdentified";

export class ColumnMetaIdentified<
  VN extends ValueName = ValueName,
> extends ColumnCommonIdentified<VN> {
  get raw(): ColumnMetaRaw<VnToCvn<VN>> {
    return new ColumnMetaRaw({
      ...this.sheetIdentifiedProps,
      colIndex: this.colIndex,
    });
  }
  get sheet(): SheetMetaIdentified {
    return new SheetMetaIdentified(this.sheetIdentifiedProps);
  }
  get primary(): ColumnIdentified<VN> {
    return new ColumnIdentified(this.columnIdentifiedProps);
  }
  get activeColumnType(): string | undefined {
    return this.raw.activeColumnType;
  }
  updateColumnType(columnType: TableColumnType): this {
    this.raw.updateColumnType(columnType);
    return this;
  }
  uniformCell<UN extends UniformRowName>(
    rowName: UN,
  ): CellIdentified<UniformRowValueName<UN>> {
    const rowIndex = this.schema.uniformRowIndex(rowName);
    const valueName = this.schema.uniformValueName(rowName);
    return new CellIdentified<UniformRowValueName<UN>>({
      ...this.columnIdentifiedProps,
      rowIndex,
      valueName,
    });
  }
}
