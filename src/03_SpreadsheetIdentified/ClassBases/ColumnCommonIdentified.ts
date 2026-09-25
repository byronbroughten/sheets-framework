import type { ValueName } from "../../01_SpreadsheetSchema/valueSchemas";
import { SheetMetaRaw } from "../../02_SpreadsheetRaw/SheetMetaRaw";
import { ColumnBaseIdentified } from "./ColumnBaseIdentified";

export abstract class ColumnCommonIdentified<
  VN extends ValueName = ValueName,
> extends ColumnBaseIdentified<VN> {
  get colIndex(): number {
    return new SheetMetaRaw(this.sheetIdentifiedProps).colIndexOfActiveColumnId(
      this.columnId,
    );
  }
}
