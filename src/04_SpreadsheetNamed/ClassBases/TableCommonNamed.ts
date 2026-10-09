import type { TableName } from "../../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { TableSchema } from "../../01_SpreadsheetSchema/configReaders/TableSchema";
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
  get tableAddress(): { tableId: string } {
    return managedTableAddress(this.schema);
  }
}
