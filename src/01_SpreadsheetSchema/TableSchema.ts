import {
  type SheetColIndex,
  SheetIndex,
  type SheetRowIndex,
} from "../00_Source/RawSource/SheetIndex";
import { Val } from "../utils/Val";
import {
  type ColumnName,
  getColumnTraitById,
  getColumnTraitByName,
  getSheetColumnNames,
  getTableColumnIds,
} from "./columnConfigsTypes";
import { ColumnSchema } from "./ColumnSchema";
import { dimensionIds } from "./dimensionIds";
import { SpreadsheetBaseSchema } from "./SpreadsheetBaseSchema";
import {
  getTableTraitByName,
  type TableConfig,
  tableConfigsByTableId,
  tableKeysByGid,
  type TableName,
} from "./tableConfigsTypes";
import { TableOrigin } from "./TableOrigin";

function tableNameFromGid(sheetGid: number): TableName {
  const byGid = tableKeysByGid();
  const [tableKey, ...otherKeys] = byGid.get(sheetGid) ?? [];
  if (tableKey === undefined) {
    throw new Error(
      `Invalid sheetGid: ${sheetGid}. Must be one of: ${[...byGid.keys()].join(", ")}`,
    );
  }
  if (otherKeys.length > 0) {
    throw new Error(
      `Sheet gid ${sheetGid} holds several managed Tables (${[tableKey, ...otherKeys].join(", ")}); reach one by its Table ID.`,
    );
  }
  return tableKey;
}

export interface TableSchemaProps<TN extends TableName> {
  sheetGid: number;
  tableName: TN;
}

export class TableSchema<
  TN extends TableName = TableName,
> extends SpreadsheetBaseSchema {
  readonly sheetGid: number;
  readonly tableName: TN;
  constructor({ sheetGid, tableName }: TableSchemaProps<TN>) {
    super();
    this.sheetGid = sheetGid;
    this.tableName = tableName;
  }
  static fromSheetName<TN extends TableName>(tableName: TN): TableSchema<TN> {
    return new TableSchema({
      tableName,
      sheetGid: getTableTraitByName(tableName, "sheetGid"),
    });
  }
  // Only for a sheet the configs record one Table on.
  static fromSheetGid(sheetGid: number): TableSchema {
    return new TableSchema({
      sheetGid,
      tableName: tableNameFromGid(sheetGid),
    });
  }
  static fromTableId(tableId: string): TableSchema {
    const { tableKey } = Val.assert(
      tableConfigsByTableId().get(tableId),
      `Table config for tableId ${tableId}`,
    );
    return TableSchema.fromSheetName(tableKey as TableName);
  }
  // With another Table the configs record on its sheet.
  get sharesSheet(): boolean {
    return (tableKeysByGid().get(this.sheetGid)?.length ?? 0) > 1;
  }
  trait<TK extends keyof TableConfig>(key: TK): TableConfig[TK] {
    return getTableTraitByName(this.tableName, key);
  }
  get tableId(): string {
    return this.trait("tableId");
  }
  get idPrefix(): string {
    return this.trait("idPrefix");
  }
  get recordedOrigin(): TableOrigin {
    return new TableOrigin({
      headerRowIndex: SheetIndex.row(this.trait("headerRowIndex")),
      startColIndex: SheetIndex.col(this.trait("startColIndex")),
    });
  }
  get recordedStartLabel(): string {
    const { headerRowIndex, startColIndex } = this.recordedOrigin;
    return this.positionLabel(headerRowIndex, startColIndex);
  }
  holdsActionCellAt(
    sheetRowIndex: SheetRowIndex,
    sheetColIndex: SheetColIndex,
  ): boolean {
    const origin = this.recordedOrigin;
    const colIndex = origin.colIndex(sheetColIndex);
    return (
      sheetRowIndex === origin.headSheetRowIndex("action") &&
      colIndex >= 0 &&
      colIndex < this.columnNames.length
    );
  }
  makeRowId(): string {
    return dimensionIds.row(this.idPrefix);
  }
  get columnIds(): MapIterator<string> {
    return getTableColumnIds(this.tableName);
  }
  get columnNames(): ColumnName<TN>[] {
    return getSheetColumnNames(this.tableName);
  }
  get nonFormulaColumnIds(): string[] {
    return [...this.columnIds].filter((columnId) => {
      return !getColumnTraitById(this.tableName, columnId, "isFormula");
    });
  }
  // Goes by the columnId index, since by name would scan the Table.
  colNameByColumnId(columnId: string): ColumnName<TN> {
    return getColumnTraitById(
      this.tableName,
      columnId,
      "columnName",
    ) as ColumnName<TN>;
  }
  columnByName<CN extends ColumnName<TN>>(
    columnName: CN,
  ): ColumnSchema<TN, CN> {
    return new ColumnSchema({
      ...this.sheetSchemaProps,
      columnName,
      columnId: getColumnTraitByName(this.tableName, columnName, "columnId"),
    });
  }
  columnById(columnId: string): ColumnSchema<TN, ColumnName<TN>> {
    return new ColumnSchema({
      ...this.sheetSchemaProps,
      columnId,
      columnName: this.colNameByColumnId(columnId),
    });
  }
  columnIdByHeader(header: string): string {
    const columnId = [...this.columnIds].find(
      (id) => getColumnTraitById(this.tableName, id, "header") === header,
    );
    if (columnId === undefined) {
      throw new Error(`"${this.tableName}" has no column headed "${header}".`);
    }
    return columnId;
  }
  columnSpecifierToStandard(
    columnSpecifier: ColumnName<TN> | ColumnName<TN>[] | "allColumns",
  ): ColumnName<TN>[] {
    if (columnSpecifier === "allColumns") {
      return this.columnNames;
    } else if (Array.isArray(columnSpecifier)) {
      return columnSpecifier;
    } else {
      return [columnSpecifier];
    }
  }
  private get sheetSchemaProps(): TableSchemaProps<TN> {
    return { sheetGid: this.sheetGid, tableName: this.tableName };
  }
}
