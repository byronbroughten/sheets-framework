import { ColumnSchema } from "../../01_SpreadsheetSchema/ColumnSchema";
import type { ValueName } from "../../01_SpreadsheetSchema/valueSchemas";
import {
  SheetBaseIdentified,
  type SheetIdentifiedProps,
} from "./SheetBaseIdentified";

export interface ColumnIdentifiedProps<
  VN extends ValueName = ValueName,
> extends SheetIdentifiedProps {
  columnId: string;
  valueName?: VN;
}

export class ColumnBaseIdentified<
  VN extends ValueName = ValueName,
> extends SheetBaseIdentified {
  readonly columnId: string;
  readonly valueName?: VN;
  constructor({ columnId, valueName, ...props }: ColumnIdentifiedProps<VN>) {
    super(props);
    this.columnId = columnId;
    this.valueName = valueName;
  }
  get schema(): ColumnSchema {
    return ColumnSchema.fromColumnId(this.sheetGid, this.columnId);
  }
  get columnIdentifiedProps(): ColumnIdentifiedProps<VN> {
    return {
      ...this.sheetIdentifiedProps,
      columnId: this.columnId,
      valueName: this.valueName,
    };
  }
}
