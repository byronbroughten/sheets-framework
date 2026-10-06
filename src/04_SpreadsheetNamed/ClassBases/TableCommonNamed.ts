import type { TableName } from "../../01_SpreadsheetSchema/sheetConfigsTypes";
import { TableSchema } from "../../01_SpreadsheetSchema/TableSchema";
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
}
