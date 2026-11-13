import { columnConfigsByName } from "../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import type { ColumnConfigStored } from "../01_SpreadsheetSchema/configReaders/makeConfigs";
import { tableConfigsByTableId } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";

// One Table's columns as the previous configs hold them; a Table new to them has none.
export class PreviousColumnConfigs {
  readonly tableKey: string | undefined;
  private readonly columnConfigsByName: Record<string, ColumnConfigStored>;
  constructor(tableId: string) {
    this.tableKey = tableConfigsByTableId().get(tableId)?.tableKey;
    if (this.tableKey === undefined) {
      this.columnConfigsByName = {};
      return;
    }
    this.columnConfigsByName = columnConfigsByName()[this.tableKey] ?? {};
  }
  get columnConfigs(): ColumnConfigStored[] {
    return Object.values(this.columnConfigsByName);
  }
  get columnIds(): string[] {
    return this.columnConfigs.map(({ columnId }) => columnId);
  }
  // Through the ID, so a renamed header keeps its previous name (ADR 0035).
  columnName(columnId: string): string | undefined {
    return Object.entries(this.columnConfigsByName).find(
      ([, columnConfig]) => columnConfig.columnId === columnId,
    )?.[0];
  }
  // A column new to the configs is required until a later run's configs say otherwise.
  emptyValueAllowed(columnId: string): boolean {
    const previous = this.columnConfigs.find(
      (columnConfig) => columnConfig.columnId === columnId,
    );
    return previous?.emptyValueAllowed ?? false;
  }
}
