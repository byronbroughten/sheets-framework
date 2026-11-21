import { type TableSchema } from "../../01_SpreadsheetSchema/configReaders/TableSchema.ts";
import {
  TableBaseIdentified,
  type TableIdentifiedProps,
} from "./TableBaseIdentified.ts";

export type RowIdentifiedProps = TableIdentifiedProps & { rowIndex: number };

export class RowBaseIdentified extends TableBaseIdentified {
  readonly rowIndex: number;
  constructor({ rowIndex, ...rest }: RowIdentifiedProps) {
    super(rest);
    this.rowIndex = rowIndex;
  }
  get schema(): TableSchema {
    return this.tableSchema;
  }
  get rowIdentifiedProps(): RowIdentifiedProps {
    return {
      rowIndex: this.rowIndex,
      ...this.tableIdentifiedProps,
    };
  }
}
