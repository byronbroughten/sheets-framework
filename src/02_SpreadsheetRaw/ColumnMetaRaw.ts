import type { CellValueName } from "../00_Source/CellValues/cellValues";
import { ColumnBaseRaw } from "./ClassBases/ColumnBaseRaw";
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
}
