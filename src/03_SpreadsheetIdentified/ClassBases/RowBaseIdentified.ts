import { SheetSchema } from "../../01_SpreadsheetSchema/SheetSchema";
import {
  SheetBaseIdentified,
  type SheetIdentifiedProps,
} from "./SheetBaseIdentified";

export interface RowIdentifiedProps extends SheetIdentifiedProps {
  rowIndex: number;
}

export class RowBaseIdentified extends SheetBaseIdentified {
  readonly rowIndex: number;
  constructor({ rowIndex, ...rest }: RowIdentifiedProps) {
    super(rest);
    this.rowIndex = rowIndex;
  }
  get schema(): SheetSchema {
    return SheetSchema.fromSheetGid(this.sheetGid);
  }
  get rowIdentifiedProps(): RowIdentifiedProps {
    return {
      rowIndex: this.rowIndex,
      ...this.sheetIdentifiedProps,
    };
  }
}
