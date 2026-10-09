import type { TableName } from "../../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { type TableSchema } from "../../01_SpreadsheetSchema/configReaders/TableSchema";
import { TableBaseIdentified } from "./TableBaseIdentified";

export abstract class TableCommonIdentified extends TableBaseIdentified {
  get schema(): TableSchema {
    return this.tableSchema;
  }
  get tableName(): TableName {
    return this.schema.tableName;
  }
}
