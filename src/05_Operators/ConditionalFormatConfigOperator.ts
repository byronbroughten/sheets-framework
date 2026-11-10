import { Val } from "@byronbroughten/utils/val";

import {
  type ColumnConfigsGeneric,
  type ConditionalFormatConfigsBase,
  makeImportLine,
} from "../01_SpreadsheetSchema/configReaders/makeConfigs";
import {
  RuleRecords,
  type RuleRecordsTab,
  type RuleRecordsTable,
} from "../01_SpreadsheetSchema/configReaders/RuleRecords";
import type { SheetRaw } from "../02_SpreadsheetRaw/SheetRaw";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import { ColumnConfigOperator } from "./ColumnConfigOperator";
import {
  indentFollowingLines,
  oneLinePerLeafFileSource,
} from "./configFileSource";
import { SpreadsheetBaseOperator } from "./SpreadsheetBaseOperator";
import { TableConfigOperator } from "./TableConfigOperator";

interface RuleTableProps {
  tableId: string;
  tableKey: string;
  columnConfigs: ColumnConfigsGeneric;
}

// Reads the rules the regen's plain property get brought into Raw, so it costs no round trip of its own.
export class ConditionalFormatConfigOperator extends SpreadsheetBaseOperator {
  static init(): ConditionalFormatConfigOperator {
    return new ConditionalFormatConfigOperator(
      SpreadsheetBaseOperator.initOperatorProps(),
    );
  }
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  get tableConfigOperator(): TableConfigOperator {
    return new TableConfigOperator(this.operatorProps);
  }
  get columnConfigOperator(): ColumnConfigOperator {
    return new ColumnConfigOperator(this.operatorProps);
  }
  newConditionalFormatConfigs(): ConditionalFormatConfigsBase {
    const tableKeysByTableId = this.tableConfigOperator.tableKeysByTableId();
    const columnConfigs = this.columnConfigOperator.newColumnConfigs();
    const { raw } = this.ss;
    const tabs = raw.activeSheetGids.map((sheetGid) =>
      this._ruleRecordsTab({
        sheet: raw.sheet(sheetGid),
        tableKeysByTableId,
        columnConfigs,
      }),
    );
    return RuleRecords.build(tabs);
  }
  private _ruleRecordsTab({
    sheet,
    tableKeysByTableId,
    columnConfigs,
  }: {
    sheet: SheetRaw;
    tableKeysByTableId: Map<string, string>;
    columnConfigs: ColumnConfigsGeneric;
  }): RuleRecordsTab {
    const tables = sheet.tableIds.flatMap((tableId) => {
      const tableKey = tableKeysByTableId.get(tableId);
      if (tableKey === undefined) return [];
      return [this._ruleRecordsTable({ tableId, tableKey, columnConfigs })];
    });
    return {
      tabName: this.schema.titleToName(sheet.title),
      sheetGid: sheet.sheetGid,
      tables,
      rules: sheet.conditionalFormatRules(),
    };
  }
  private _ruleRecordsTable({
    tableId,
    tableKey,
    columnConfigs,
  }: RuleTableProps): RuleRecordsTable {
    const table = this.ss.raw.table(tableId);
    const columnNamesById = columnNamesByColumnId(
      Val.assert(columnConfigs[tableKey], `column configs of ${tableKey}`),
    );
    const columnNamesLeftToRight = table.fullTableColIndexes.map((colIndex) =>
      Val.assert(
        columnNamesById.get(table.columnResolver.columnIdAt(colIndex)),
        `column name at column index ${colIndex} of ${tableKey}`,
      ),
    );
    return {
      tableKey,
      tableId,
      layout: {
        headerTopLeft: {
          rowIndex: table.startRowIndex,
          colIndex: table.startColumnIndex,
        },
        columnNamesLeftToRight,
      },
      endRowIndex: table.endRowIndex,
    };
  }
  toFileSource(): string {
    const { tables, tabs } = this.newConditionalFormatConfigs();
    return [
      makeImportLine(["makeConditionalFormatConfigs"]),
      ``,
      `export const conditionalFormatConfigs = makeConditionalFormatConfigs({`,
      `  "tables": ${indentFollowingLines(oneLinePerLeafFileSource(tables))},`,
      `  "tabs": ${indentFollowingLines(oneLinePerLeafFileSource(tabs))}`,
      `});`,
      ``,
    ].join("\n");
  }
}

function columnNamesByColumnId(
  tableColumnConfigs: ColumnConfigsGeneric[string],
): Map<string, string> {
  return Object.entries(tableColumnConfigs).reduce(
    (namesById, [columnName, { columnId }]) =>
      namesById.set(columnId, columnName),
    new Map<string, string>(),
  );
}
