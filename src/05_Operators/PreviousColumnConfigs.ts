import { columnConfigsByName } from "../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import type { ColumnConfigStored } from "../01_SpreadsheetSchema/configReaders/makeConfigs";
import { tableConfigsByTableId } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";

// One Table's columns as the previous configs hold them; a Table new to them has none.
export class PreviousColumnConfigs {
  readonly columnConfigs: ColumnConfigStored[];
  constructor(tableId: string) {
    const tableKey = tableConfigsByTableId().get(tableId)?.tableKey;
    if (tableKey === undefined) {
      this.columnConfigs = [];
      return;
    }
    this.columnConfigs = Object.values(columnConfigsByName()[tableKey] ?? {});
  }
  get columnIds(): string[] {
    return this.columnConfigs.map(({ columnId }) => columnId);
  }
  // A column new to the configs is required until a later run's configs say otherwise.
  emptyValueAllowed(columnId: string): boolean {
    const previous = this.columnConfigs.find(
      (columnConfig) => columnConfig.columnId === columnId,
    );
    return previous?.emptyValueAllowed ?? false;
  }
}
