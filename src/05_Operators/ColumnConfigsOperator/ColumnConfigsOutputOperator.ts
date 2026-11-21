import { Val } from "@byronbroughten/utils/val";

import { columnNameOfHeader } from "../../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import {
  type ColumnConfigsGeneric,
  makeImportLine,
} from "../../01_SpreadsheetSchema/configReaders/makeConfigs";
import type { ValueName } from "../../01_SpreadsheetSchema/configReaders/valueSchemas";
import { ColumnConfigsBaseOperator } from "../ClassBases/ColumnConfigsBaseOperator";
import { oneLinePerLeafFileSource } from "../configFileSource";
import { PreviousColumnConfigs } from "../PreviousColumnConfigs";
import { ColumnValueSourcesOperator } from "./ColumnValueSourcesOperator";

export interface ApiTableColumns {
  tableKey: string;
  columnConfigs: ColumnConfigsGeneric[string];
}

export class ColumnConfigsOutputOperator extends ColumnConfigsBaseOperator {
  get valueSourcesOperator(): ColumnValueSourcesOperator {
    return new ColumnValueSourcesOperator(this.operatorProps);
  }
  newColumnConfigs(): ColumnConfigsGeneric {
    const tableKeysByTableId = this.tableConfigOperator.tableKeysByTableId();
    const { apiColumnIdentities, valueSourcesOperator } = this;
    const columnConfigs: ColumnConfigsGeneric = {};
    for (const identity of apiColumnIdentities.columnIdentities()) {
      const { tableId, columnId } = identity;
      const describedColumn = apiColumnIdentities.describedColumn(identity);
      const { header } = describedColumn;
      const tableKey = Val.assert(
        tableKeysByTableId.get(tableId),
        "API-access Table key",
      );
      const columnName = columnNameOfHeader(header);
      if (!columnConfigs[tableKey]) {
        columnConfigs[tableKey] = {};
      }
      const tableColumnConfigs = columnConfigs[tableKey];
      if (tableColumnConfigs[columnName]) {
        throw new Error(
          `generateColumnConfigFileSource: duplicate column name "${columnName}" ` +
            `derived from header "${header}" on Table "${tableKey}".`,
        );
      }
      const valueName =
        valueSourcesOperator._valueSource(describedColumn, tableKeysByTableId)
          ?.valueName ?? this.schema.titleToName(describedColumn.valueTitle());
      tableColumnConfigs[columnName] = {
        columnId,
        header,
        valueName: valueName as ValueName,
        isFormula: describedColumn.isFormula,
        emptyValueAllowed:
          this.columnConfigSync.emptyValueAllowedByIds
            .get(tableId)
            ?.get(columnId) ??
          new PreviousColumnConfigs(tableId).emptyValueAllowed(columnId),
        customDefaultValue: null,
        formula: describedColumn.topFormula ?? null,
        columnType: {
          type: describedColumn.columnType ?? null,
          rule: describedColumn.validationRule ?? null,
        },
        actionRowCell: describedColumn.actionRowCell ?? null,
        groupHeading1: describedColumn.groupHeading1 || null,
      };
    }
    return this.installedBraceConfigs.inBraceForm({
      tableConfigs: this.tableConfigOperator.newTableConfigs(),
      columnConfigs,
    });
  }
  apiTableColumnsByTableId(): Map<string, ApiTableColumns> {
    const byTableId = this.identifiedApiTableColumnsByTableId();
    const tableKeysByTableId = this.tableConfigOperator.tableKeysByTableId();
    for (const [tableId, tableKey] of tableKeysByTableId) {
      Val.assert(byTableId.get(tableId), `column configs of ${tableKey}`);
    }
    return byTableId;
  }
  // Read before a send, an API-access Table whose column IDs the send mints has no column configs yet, so it is left out.
  identifiedApiTableColumnsByTableId(): Map<string, ApiTableColumns> {
    const columnConfigs = this.newColumnConfigs();
    const tableKeysByTableId = this.tableConfigOperator.tableKeysByTableId();
    return [...tableKeysByTableId].reduce((byTableId, [tableId, tableKey]) => {
      const tableColumnConfigs = columnConfigs[tableKey];
      if (tableColumnConfigs === undefined) return byTableId;
      return byTableId.set(tableId, {
        tableKey,
        columnConfigs: tableColumnConfigs,
      });
    }, new Map<string, ApiTableColumns>());
  }
  toFileSource(): string {
    return [
      `${makeImportLine(["makeColumnConfigs"])}`,
      ``,
      `export const columnConfigs = makeColumnConfigs(${oneLinePerLeafFileSource(
        this.newColumnConfigs(),
      )});`,
      ``,
    ].join("\n");
  }
}
