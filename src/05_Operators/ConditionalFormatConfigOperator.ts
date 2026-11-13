import { Val } from "@byronbroughten/utils/val";

import {
  type ConditionalFormatConfigsBase,
  makeImportLine,
} from "../01_SpreadsheetSchema/configReaders/makeConfigs";
import {
  RuleRecords,
  type RuleRecordsTab,
  type RuleRecordsTable,
} from "../01_SpreadsheetSchema/configReaders/RuleRecords";
import type { SheetRaw } from "../02_SpreadsheetRaw/SheetRaw";
import type { TableRaw } from "../02_SpreadsheetRaw/TableRaw";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import {
  type ApiTableColumns,
  ColumnConfigsOperator,
} from "./ColumnConfigsOperator";
import {
  columnNamesByColumnId,
  oneLinePerLeafFileSource,
  tablesAndTabsFileSource,
} from "./configFileSource";
import { SpreadsheetBaseOperator } from "./SpreadsheetBaseOperator";

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
  get columnConfigOperator(): ColumnConfigsOperator {
    return new ColumnConfigsOperator(this.operatorProps);
  }
  newConditionalFormatConfigs(): ConditionalFormatConfigsBase {
    const apiTableColumnsByTableId =
      this.columnConfigOperator.apiTableColumnsByTableId();
    const { raw } = this.ss;
    const tabs = raw.activeSheetGids.map((sheetGid) =>
      this._ruleRecordsTab(raw.sheet(sheetGid), apiTableColumnsByTableId),
    );
    return RuleRecords.build(tabs);
  }
  private _ruleRecordsTab(
    sheet: SheetRaw,
    apiTableColumnsByTableId: Map<string, ApiTableColumns>,
  ): RuleRecordsTab {
    const tables = sheet.tableIds.flatMap((tableId) => {
      const apiTableColumns = apiTableColumnsByTableId.get(tableId);
      if (apiTableColumns === undefined) return [];
      return [this._ruleRecordsTable(tableId, apiTableColumns)];
    });
    return {
      tabName: this.schema.titleToName(sheet.title),
      sheetGid: sheet.sheetGid,
      tables,
      rules: sheet.conditionalFormatRules(),
    };
  }
  private _ruleRecordsTable(
    tableId: string,
    { tableKey, columnConfigs }: ApiTableColumns,
  ): RuleRecordsTable {
    const table = this.ss.raw.table(tableId);
    const columnNamesById = columnNamesByColumnId(columnConfigs);
    const columnNamesLeftToRight = table.fullTableColIndexes.map((colIndex) =>
      columnNameAt({ table, colIndex, columnNamesById }),
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
    return [
      makeImportLine(["makeConditionalFormatConfigs"]),
      ``,
      `export const conditionalFormatConfigs = makeConditionalFormatConfigs(${tablesAndTabsFileSource(
        this.newConditionalFormatConfigs(),
        oneLinePerLeafFileSource,
      )});`,
      ``,
    ].join("\n");
  }
}

interface ColumnNameAtProps {
  table: TableRaw;
  colIndex: number;
  columnNamesById: Map<string, string>;
}

// Undefined for a column with no ID: a run that mints none leaves it out of the configs.
function columnNameAt({
  table,
  colIndex,
  columnNamesById,
}: ColumnNameAtProps): string | undefined {
  const columnId = table.columnResolver.columnIdAt(colIndex);
  if (columnId === "") return undefined;
  return Val.assert(
    columnNamesById.get(columnId),
    `column name at column index ${colIndex} of ${table.name}`,
  );
}
