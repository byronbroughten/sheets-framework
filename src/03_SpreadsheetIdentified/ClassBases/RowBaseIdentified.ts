import { SheetSchema } from "../../01_SpreadsheetSchema/SheetSchema";
import {
  TableBaseIdentified,
  type TableIdentifiedProps,
} from "./TableBaseIdentified";

export type RowIdentifiedProps = TableIdentifiedProps & { rowIndex: number };

export class RowBaseIdentified extends TableBaseIdentified {
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
      ...this.tableIdentifiedProps,
    };
  }
}
