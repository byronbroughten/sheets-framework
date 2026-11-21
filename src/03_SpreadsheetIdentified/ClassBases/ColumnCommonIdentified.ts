import type { ValueName } from "../../01_SpreadsheetSchema/configReaders/valueSchemas.ts";
import { TableColumnResolverRaw } from "../../02_SpreadsheetRaw/TableRaw/TableColumnResolverRaw.ts";
import { ColumnBaseIdentified } from "./ColumnBaseIdentified.ts";

export abstract class ColumnCommonIdentified<
  VN extends ValueName = ValueName,
> extends ColumnBaseIdentified<VN> {
  get colIndex(): number {
    return new TableColumnResolverRaw(this.tableIdentifiedProps).colIndexOf(
      this.columnId,
    );
  }
}
