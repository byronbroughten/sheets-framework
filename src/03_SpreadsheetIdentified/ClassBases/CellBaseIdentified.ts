import type { ValueName } from "../../01_SpreadsheetSchema/configReaders/valueSchemas";
import {
  ColumnBaseIdentified,
  type ColumnIdentifiedProps,
} from "./ColumnBaseIdentified";

export type CellIdentifiedProps<VN extends ValueName = ValueName> =
  ColumnIdentifiedProps<VN> & { rowIndex: number };

export class CellBaseIdentified<
  VN extends ValueName = ValueName,
> extends ColumnBaseIdentified<VN> {
  readonly rowIndex: number;
  constructor({ rowIndex, ...props }: CellIdentifiedProps<VN>) {
    super(props);
    this.rowIndex = rowIndex;
  }
  get cellIdentifiedProps(): CellIdentifiedProps<VN> {
    return {
      rowIndex: this.rowIndex,
      ...this.columnIdentifiedProps,
    };
  }
}
