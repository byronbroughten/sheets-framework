import { ColumnConfigsBaseOperator } from "../ClassBases/ColumnConfigsBaseOperator";
import type { UntypedHeadersByTableName } from "../SpreadsheetBaseOperator";
import { ColumnValueSourcesOperator } from "./ColumnValueSourcesOperator";

export class UntypedColumnsOperator extends ColumnConfigsBaseOperator {
  get valueSourcesOperator(): ColumnValueSourcesOperator {
    return new ColumnValueSourcesOperator(this.operatorProps);
  }
  private get untypedHeadersByTableName(): UntypedHeadersByTableName {
    return this.columnConfigSync.untypedHeadersByTableName;
  }
  assertSyncedToSpreadsheet(): void {
    if (!this.columnConfigSync.syncedToSpreadsheet) {
      throw new Error(
        "ColumnConfigsOperator has not yet synced to the spreadsheet.",
      );
    }
  }
  recordUntypedColumns(): this {
    this._recordUntypedColumns();
    this._logUntypedColumns();
    this.columnConfigSync.syncedToSpreadsheet = true;
    return this;
  }
  private _recordUntypedColumns(): void {
    this.untypedHeadersByTableName.clear();
    const { apiColumnIdentities, valueSourcesOperator } = this;
    for (const identity of apiColumnIdentities.columnIdentities()) {
      const describedColumn = apiColumnIdentities.describedColumn(identity);
      if (
        valueSourcesOperator._sourceColumn(describedColumn) === undefined &&
        describedColumn.declaredValueTitle() === undefined
      ) {
        this._recordUntypedColumn(
          this.ss.raw.table(identity.tableId).name,
          describedColumn.header,
        );
      }
    }
  }
  private _recordUntypedColumn(tableName: string, header: string): void {
    const headers = this.untypedHeadersByTableName.get(tableName) ?? [];
    headers.push(header);
    this.untypedHeadersByTableName.set(tableName, headers);
  }
  private _logUntypedColumns(): void {
    for (const [tableName, headers] of this.untypedHeadersByTableName) {
      Logger.log(
        `Untyped columns on Table "${tableName}" (${headers.length}): ${headers.join(", ")}`,
      );
    }
  }
  // Undefined, not "", so a fully typed spreadsheet still reports "Succeeded".
  untypedColumnsSummary(): string | undefined {
    this.assertSyncedToSpreadsheet();
    const untypedHeaders = Array.from(this.untypedHeadersByTableName.values());
    if (untypedHeaders.length === 0) {
      return undefined;
    }
    const columnCount = untypedHeaders.reduce(
      (count, headers) => count + headers.length,
      0,
    );
    const sentences = [
      `Succeeded, but ${columnCount} column(s) across ${untypedHeaders.length} ` +
        `Table(s) are untyped, so their value names were guessed. See the ` +
        `execution log for the list.`,
    ];
    const blankSampleNames = this._blankSampleTableNames();
    if (blankSampleNames.length > 0) {
      const names = blankSampleNames.map((name) => `"${name}"`).join(", ");
      sentences.push(
        `On ${blankSampleNames.length} of those Table(s) the top data row ` +
          `was blank, so the guess had no sample behind it: ${names}.`,
      );
    }
    return sentences.join(" ");
  }
  // Derived, since blank facts are deliberately indistinguishable from real ones.
  private _blankSampleTableNames(): string[] {
    const names: string[] = [];
    for (const tableId of this.apiColumnIdentities.tableIdsApiAccesses) {
      const table = this.ss.raw.table(tableId);
      if (!this.untypedHeadersByTableName.has(table.name)) continue;
      if (!table.topDataRowIsBlank()) continue;
      names.push(table.name);
    }
    return names;
  }
}
