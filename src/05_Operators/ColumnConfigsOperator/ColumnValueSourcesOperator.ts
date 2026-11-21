import { Val } from "@byronbroughten/utils/val";

import { isFrameworkValueName } from "../../00_Source/CellValues/frameworkValueSchemas";
import { columnNameOfHeader } from "../../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import type {
  ColumnProfileRaw,
  TableColumnReference,
} from "../../02_SpreadsheetRaw/ColumnProfileRaw";
import { ColumnConfigsBaseOperator } from "../ClassBases/ColumnConfigsBaseOperator";
import {
  type SourceColumn,
  sourceLabel,
  type ValueSource,
} from "../SpreadsheetBaseOperator";

export class ColumnValueSourcesOperator extends ColumnConfigsBaseOperator {
  valueSources(): ValueSource[] {
    const tableKeysByTableId = this.tableConfigOperator.tableKeysByTableId();
    const { apiColumnIdentities } = this;
    const sourcesByValueName = new Map<string, ValueSource>();
    for (const identity of apiColumnIdentities.columnIdentities()) {
      const source = this._valueSource(
        apiColumnIdentities.describedColumn(identity),
        tableKeysByTableId,
      );
      if (source === undefined) continue;
      const claimed = sourcesByValueName.get(source.valueName);
      if (claimed !== undefined && !isSameSourceColumn(claimed, source)) {
        throw new Error(
          `Value sources ${this._sourceLabel(claimed)} and ${this._sourceLabel(source)} ` +
            `both take the value name "${source.valueName}". Rename a Table or a header.`,
        );
      }
      sourcesByValueName.set(source.valueName, source);
    }
    return [...sourcesByValueName.values()];
  }
  _valueSource(
    describedColumn: ColumnProfileRaw,
    tableKeysByTableId: Map<string, string>,
  ): ValueSource | undefined {
    const sourceColumn = this._sourceColumn(describedColumn);
    if (sourceColumn === undefined) return undefined;
    const { tableId, header } = sourceColumn;
    return {
      ...sourceColumn,
      valueName: valueSourceName({
        tableKey: Val.assert(
          tableKeysByTableId.get(tableId),
          "source Table key",
        ),
        reference: { tableName: this.ss.raw.table(tableId).name, header },
      }),
    };
  }
  _sourceColumn(describedColumn: ColumnProfileRaw): SourceColumn | undefined {
    const reference = describedColumn.validationColumnReference();
    if (reference === undefined || describedColumn.isIdColumn) return undefined;
    const tableId = [...this.apiColumnIdentities.tableIdsApiAccesses].find(
      (apiTableId) =>
        this.ss.raw.table(apiTableId).name === reference.tableName,
    );
    if (tableId === undefined) return undefined;
    if (!this.ss.raw.table(tableId).profile.hasHeader(reference.header)) {
      return undefined;
    }
    return { tableId, header: reference.header };
  }
  private _sourceLabel({ tableId, header }: SourceColumn): string {
    return sourceLabel({ tableName: this.ss.raw.table(tableId).name, header });
  }
}

function isSameSourceColumn(a: SourceColumn, b: SourceColumn): boolean {
  return a.tableId === b.tableId && a.header === b.header;
}

function valueSourceName({
  tableKey,
  reference,
}: {
  tableKey: string;
  reference: TableColumnReference;
}): string {
  const headerName = columnNameOfHeader(reference.header);
  const valueName =
    tableKey + headerName.charAt(0).toUpperCase() + headerName.slice(1);
  if (isFrameworkValueName(valueName)) {
    throw new Error(
      `Value source ${sourceLabel(reference)} takes the value name "${valueName}", ` +
        `which is already the framework value name "${valueName}". Rename the Table or the header.`,
    );
  }
  return valueName;
}
