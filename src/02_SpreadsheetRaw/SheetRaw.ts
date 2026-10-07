import { SpreadsheetSchema } from "../01_SpreadsheetSchema/SpreadsheetSchema";
import {
  SpreadsheetBaseRaw,
  type SpreadsheetRawProps,
} from "./ClassBases/SpreadsheetBaseRaw";

export interface SheetRawProps extends SpreadsheetRawProps {
  sheetGid: number;
}

export class SheetRaw extends SpreadsheetBaseRaw {
  readonly sheetGid: number;
  constructor({ sheetGid, ...props }: SheetRawProps) {
    super(props);
    this.sheetGid = sheetGid;
  }
  get schema(): SpreadsheetSchema {
    return new SpreadsheetSchema();
  }
  get title(): string {
    const title = this.sheetsStateRaw.get(this.sheetGid)?.working.title;
    if (title === undefined) {
      throw new Error(
        `Sheet title is null for sheetGid ${this.sheetGid}. Ensure that the sheet properties have been fetched.`,
      );
    }
    return title;
  }
  // The managed Tables only: the configs place them, so a tab's unmanaged Tables aren't listed.
  get tableIds(): string[] {
    return this.schema
      .tablesOnGid(this.sheetGid)
      .map((table) => table.tableId);
  }
}
