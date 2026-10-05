import { TableSchema } from "../../01_SpreadsheetSchema/TableSchema";
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
  get schema(): TableSchema {
    return TableSchema.fromSheetGid(this.sheetGid);
  }
  get rowIdentifiedProps(): RowIdentifiedProps {
    return {
      rowIndex: this.rowIndex,
      ...this.tableIdentifiedProps,
    };
  }
}
