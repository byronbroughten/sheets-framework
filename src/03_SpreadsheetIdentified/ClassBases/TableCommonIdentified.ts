import type { TableName } from "../../01_SpreadsheetSchema/configReaders/tableConfigsTypes.ts";
import { type TableSchema } from "../../01_SpreadsheetSchema/configReaders/TableSchema.ts";
import { TableBaseIdentified } from "./TableBaseIdentified.ts";

export abstract class TableCommonIdentified extends TableBaseIdentified {
  get schema(): TableSchema {
    return this.tableSchema;
  }
  get tableName(): TableName {
    return this.schema.tableName;
  }
}
