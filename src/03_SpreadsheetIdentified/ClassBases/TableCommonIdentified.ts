import type { TableName } from "../../01_SpreadsheetSchema/tableConfigsTypes";
import { TableSchema } from "../../01_SpreadsheetSchema/TableSchema";
import { TableBaseIdentified } from "./TableBaseIdentified";

export abstract class TableCommonIdentified extends TableBaseIdentified {
  get schema(): TableSchema {
    return TableSchema.fromSheetGid(this.sheetGid);
  }
  get tableName(): TableName {
    return this.schema.tableName;
  }
}
