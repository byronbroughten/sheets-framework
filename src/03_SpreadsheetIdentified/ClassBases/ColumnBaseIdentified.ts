import type { ColumnSchema } from "../../01_SpreadsheetSchema/configReaders/ColumnSchema";
import type { ValueName } from "../../01_SpreadsheetSchema/configReaders/valueSchemas";
import {
  TableBaseIdentified,
  type TableIdentifiedProps,
} from "./TableBaseIdentified";

export type ColumnIdentifiedProps<VN extends ValueName = ValueName> =
  TableIdentifiedProps & { columnId: string; valueName?: VN };

export class ColumnBaseIdentified<
  VN extends ValueName = ValueName,
> extends TableBaseIdentified {
  readonly columnId: string;
  readonly valueName?: VN;
  constructor({ columnId, valueName, ...props }: ColumnIdentifiedProps<VN>) {
    super(props);
    this.columnId = columnId;
    this.valueName = valueName;
  }
  get schema(): ColumnSchema {
    return this.tableSchema.columnById(this.columnId);
  }
  get columnIdentifiedProps(): ColumnIdentifiedProps<VN> {
    return {
      ...this.tableIdentifiedProps,
      columnId: this.columnId,
      valueName: this.valueName,
    };
  }
}
