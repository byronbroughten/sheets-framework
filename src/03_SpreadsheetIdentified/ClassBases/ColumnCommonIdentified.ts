import type { ValueName } from "../../01_SpreadsheetSchema/configReaders/valueSchemas";
import { TableColumnResolverRaw } from "../../02_SpreadsheetRaw/TableRaw/TableColumnResolverRaw";
import { ColumnBaseIdentified } from "./ColumnBaseIdentified";

export abstract class ColumnCommonIdentified<
  VN extends ValueName = ValueName,
> extends ColumnBaseIdentified<VN> {
  get colIndex(): number {
    return new TableColumnResolverRaw(this.tableIdentifiedProps).colIndexOf(
      this.columnId,
    );
  }
}
