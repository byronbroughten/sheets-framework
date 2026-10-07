import type { CellValueName } from "../00_Source/CellValues/cellValues";
import type { TableColumnType } from "../00_Source/RawSource/RawSource";
import { ColumnBaseRaw } from "./ClassBases/ColumnBaseRaw";
import type { TableEndColumnHeadCells } from "./ClassTypes/StateRaw";
import { ColumnRaw } from "./ColumnRaw";
import { SheetMetaRaw } from "./SheetMetaRaw";

export class ColumnMetaRaw<
  VN extends CellValueName = CellValueName,
> extends ColumnBaseRaw {
  get table(): SheetMetaRaw {
    return new SheetMetaRaw(this.tableRawProps);
  }
  get primary(): ColumnRaw<VN> {
    return new ColumnRaw<VN>(this.columnRawProps);
  }
  updateColumnType(columnType: TableColumnType): this {
    this.table.assertTableIsKnown();
    this.table.queueTableWrite({
      action: "updateColumnType",
      colIndex: this.colIndex,
      columnType,
    });
    this._ensureColumnState(this.colIndex).columnType = columnType;
    return this;
  }
  initHeadCells({
    columnId,
    header,
    groupHeading1,
  }: TableEndColumnHeadCells): this {
    const { primary } = this;
    primary.headCell("columnId").updateValue(columnId);
    primary.headCell("header").updateValue(header);
    if (groupHeading1 !== undefined) {
      primary.headCell("groupHeading1").updateValue(groupHeading1);
    }
    return this;
  }
}
