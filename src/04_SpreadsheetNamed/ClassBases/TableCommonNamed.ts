import type { TableName } from "../../01_SpreadsheetSchema/tableConfigsTypes";
import { TableSchema } from "../../01_SpreadsheetSchema/TableSchema";
import type { TableAddressRaw } from "../../02_SpreadsheetRaw/ClassBases/TableBaseRaw";
import { managedTableAddress } from "../../03_SpreadsheetIdentified/ClassBases/TableBaseIdentified";
import { TableBaseNamed } from "./TableBaseNamed";

export abstract class TableCommonNamed<
  TN extends TableName,
> extends TableBaseNamed<TN> {
  get schema(): TableSchema<TN> {
    return TableSchema.fromSheetName(this.tableName);
  }
  get sheetGid(): number {
    return this.schema.sheetGid;
  }
  get tableAddress(): TableAddressRaw {
    return managedTableAddress(this.schema);
  }
}
