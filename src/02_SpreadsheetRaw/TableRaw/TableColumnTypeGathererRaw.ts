import type {
  TableColumnPropertiesUpdate,
  TableColumnSnapshot,
  TableColumnValidationRule,
} from "../../00_Source/RawSource/RawSource";
import { TableBaseRaw, tableColumnLabel } from "../ClassBases/TableBaseRaw";
import { queuedColumnType } from "../ClassTypes/queuedColumnType";
import type { TableWrites } from "../ClassTypes/StateRaw";
import { TableRaw } from "../TableRaw";

type ColumnTypes = TableWrites["columnTypes"];

export class TableColumnTypeGathererRaw extends TableBaseRaw {
  get table(): TableRaw {
    return new TableRaw(this.tableRawProps);
  }
  // A column type queued this flush is sent before growth, so it is the one the new rows meet.
  columnTypesAfterFlush(): Map<number, string | undefined> {
    const { table } = this;
    const columnTypes = new Map(
      table.columnProperties.map(({ columnIndex, columnType }) => [
        columnIndex,
        columnType,
      ]),
    );
    const unfetched = table.fullTableColIndexes.find(
      (colIndex) => !columnTypes.has(colIndex),
    );
    if (unfetched !== undefined) {
      throw new Error(
        `${table.tableLabel} has no fetched type for ${tableColumnLabel(unfetched)}; refetch it before appending a row.`,
      );
    }
    table.writes.columnTypes.forEach((columnType, colIndex) =>
      columnTypes.set(colIndex, queuedColumnType.typeOf(columnType)),
    );
    return columnTypes;
  }
  // Google resets every dropdown's style and colours on a columnProperties send; observed: docs/sheets-behaviour.md#creating-and-updating-tables.
  // Counted as sent, so a dropdown the send removes, which has no style to keep, isn't.
  assertDropdownStyleResetAllowed(): void {
    const { columnTypes } = this.table.writes;
    if (columnTypes.size === 0) return;
    if (!this.spreadsheetStateRaw.rawSource.refusesDropdownStyleReset) return;
    const dropdowns = this._columnTypesColumnProperties(columnTypes).filter(
      ({ dataValidationRule }) => dataValidationRule !== undefined,
    );
    if (dropdowns.length === 0) return;
    throw new Error(
      `Refusing to set column types on ${this._tableLabel()} under dropdownStyleReset: "ask": it would reset the style and colours of its dropdowns ${dropdowns.map(({ columnName }) => columnName).join(", ")}.`,
    );
  }
  gatherSetTableColumnPropertiesOperation(): void {
    const { table } = this;
    const { columnTypes } = table.writes;
    if (columnTypes.size === 0) return;
    this._assertColumnTypesUpdateAllowed(columnTypes);
    table.writeOperations.setTableColumnProperties.push({
      kind: "setTableColumnProperties",
      tableId: table.tableId,
      columnProperties: this._columnTypesColumnProperties(columnTypes),
    });
  }
  private _assertColumnTypesUpdateAllowed(columnTypes: ColumnTypes): void {
    const { table } = this;
    const tableLabel = this._tableLabel();
    const fetched = table.columnProperties;
    if (fetched.length === 0) {
      throw new Error(
        `${tableLabel} has no fetched column properties; refetch it before setting a column type.`,
      );
    }
    if (
      table.writeOperations.insertTableEndColumns.some(
        ({ newColumns }) => newColumns.sheetId === this.sheetGid,
      ) ||
      table.writeOperations.changeTableColumns.some(
        ({ sheetId }) => sheetId === this.sheetGid,
      )
    ) {
      throw new Error(
        `Refusing to set column types on ${tableLabel}: the same flush inserts or deletes a column on that sheet.`,
      );
    }
    columnTypes.forEach((_, columnIndex) => {
      if (!fetched.some((column) => column.columnIndex === columnIndex)) {
        throw new Error(
          `${tableLabel} has no fetched ${tableColumnLabel(columnIndex)}.`,
        );
      }
    });
  }
  // Full list with each rule, since a partial columnProperties replaces the rest and a dropdown sent without its rule is refused.
  private _columnTypesColumnProperties(
    columnTypes: ColumnTypes,
  ): TableColumnPropertiesUpdate[] {
    return this.table.columnProperties.map((column) => {
      const { columnIndex } = column;
      if (column.columnName === undefined) {
        throw new Error(
          `${this._tableLabel()} ${tableColumnLabel(columnIndex)} has no columnName; refusing to replace column properties.`,
        );
      }
      const queued = columnTypes.get(columnIndex);
      const baseline =
        queued === undefined ? undefined : queuedColumnType.baselineOf(queued);
      if (baseline !== undefined) {
        return {
          columnIndex,
          columnName: column.columnName,
          ...queuedColumnType.properties(baseline),
        };
      }
      const columnType =
        queued === undefined
          ? column.columnType
          : queuedColumnType.typeOf(queued);
      const dataValidationRule = validationRuleOf(column);
      return {
        columnIndex,
        columnName: column.columnName,
        ...(columnType !== undefined ? { columnType } : {}),
        ...(dataValidationRule !== undefined ? { dataValidationRule } : {}),
      };
    });
  }
  private _tableLabel(): string {
    return `Table ${this.table.tableId} on "${this.table.sheet.title}"`;
  }
}

// Google reads every rule back with its condition type, so a column with none has no rule to resend.
function validationRuleOf(
  column: TableColumnSnapshot,
): TableColumnValidationRule | undefined {
  const { dataValidationConditionType, dataValidationValues } = column;
  if (dataValidationConditionType === undefined) return undefined;
  return {
    condition: dataValidationConditionType,
    values: dataValidationValues,
  };
}
