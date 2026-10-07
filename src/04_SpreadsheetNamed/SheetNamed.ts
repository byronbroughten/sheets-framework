import { SpreadsheetSchema } from "../01_SpreadsheetSchema/SpreadsheetSchema";
import type { TableName } from "../01_SpreadsheetSchema/tableConfigsTypes";
import { SheetRaw } from "../02_SpreadsheetRaw/SheetRaw";
import {
  SpreadsheetBaseNamed,
  type SpreadsheetNamedProps,
} from "./ClassBases/SpreadsheetBaseNamed";

export interface SheetNamedProps extends SpreadsheetNamedProps {
  sheetGid: number;
}

export class SheetNamed extends SpreadsheetBaseNamed {
  readonly sheetGid: number;
  constructor({ sheetGid, ...props }: SheetNamedProps) {
    super(props);
    this.sheetGid = sheetGid;
  }
  get schema(): SpreadsheetSchema {
    return new SpreadsheetSchema();
  }
  get raw(): SheetRaw {
    return new SheetRaw({
      ...this.spreadsheetRawProps,
      sheetGid: this.sheetGid,
    });
  }
  get title(): string {
    return this.raw.title;
  }
  get tableNames(): TableName[] {
    return this.schema
      .tablesOnGid(this.sheetGid)
      .map((table) => table.tableName);
  }
}
