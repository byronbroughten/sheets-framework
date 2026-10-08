import type { TableName } from "../../01_SpreadsheetSchema/tableConfigsTypes";
import { type TableSchema } from "../../01_SpreadsheetSchema/TableSchema";
import { TableBaseIdentified } from "./TableBaseIdentified";

export abstract class TableCommonIdentified extends TableBaseIdentified {
  get schema(): TableSchema {
    return this.tableSchema;
  }
  get tableName(): TableName {
    return this.schema.tableName;
  }
}
