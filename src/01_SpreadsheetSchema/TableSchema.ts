import {
  type ColumnName,
  getColumnTraitById,
  getColumnTraitByName,
  getSheetColumnIds,
  getSheetColumnNames,
} from "./columnConfigsTypes";
import { ColumnSchema } from "./ColumnSchema";
import { dimensionIds } from "./dimensionIds";
import { SpreadsheetBaseSchema } from "./SpreadsheetBaseSchema";
import {
  getTableTraitByGid,
  getTableTraitByName,
  type TableConfig,
  tableConfigsByGid,
  type TableName,
} from "./tableConfigsTypes";

function sheetNameFromGid(sheetGid: number): TableName {
  const byGid = tableConfigsByGid();
  if (!byGid.has(sheetGid)) {
    throw new Error(
      `Invalid sheetGid: ${sheetGid}. Must be one of: ${[...byGid.keys()].join(", ")}`,
    );
  }
  return getTableTraitByGid(sheetGid, "tableKey") as TableName;
}

export interface TableSchemaProps<TN extends TableName> {
  sheetGid: number;
  sheetName: TN;
}

export class TableSchema<
  TN extends TableName = TableName,
> extends SpreadsheetBaseSchema {
  readonly sheetGid: number;
  readonly sheetName: TN;
  constructor({ sheetGid, sheetName }: TableSchemaProps<TN>) {
    super();
    this.sheetGid = sheetGid;
    this.sheetName = sheetName;
  }
  static fromSheetName<TN extends TableName>(sheetName: TN): TableSchema<TN> {
    return new TableSchema({
      sheetName,
      sheetGid: getTableTraitByName(sheetName, "sheetGid"),
    });
  }
  static fromSheetGid(sheetGid: number): TableSchema {
    return new TableSchema({
      sheetGid,
      sheetName: sheetNameFromGid(sheetGid),
    });
  }
  trait<TK extends keyof TableConfig>(key: TK): TableConfig[TK] {
    return getTableTraitByGid(this.sheetGid, key);
  }
  get tableId(): string {
    return this.trait("tableId");
  }
  get idPrefix(): string {
    return this.trait("idPrefix");
  }
  makeRowId(): string {
    return dimensionIds.row(this.idPrefix);
  }
  get columnIds(): MapIterator<string> {
    return getSheetColumnIds(this.sheetGid);
  }
  get columnNames(): ColumnName<TN>[] {
    return getSheetColumnNames(this.sheetName);
  }
  get nonFormulaColumnIds(): string[] {
    return [...this.columnIds].filter((columnId) => {
      return !getColumnTraitById(this.sheetGid, columnId, "isFormula");
    });
  }
  // Goes by gid, the only O(1) columnId -> columnName index; by name would scan the sheet.
  colNameByColumnId(columnId: string): ColumnName<TN> {
    return getColumnTraitById(
      this.sheetGid,
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
      columnId: getColumnTraitByName(this.sheetName, columnName, "columnId"),
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
      (id) => getColumnTraitById(this.sheetGid, id, "header") === header,
    );
    if (columnId === undefined) {
      throw new Error(`"${this.sheetName}" has no column headed "${header}".`);
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
    return { sheetGid: this.sheetGid, sheetName: this.sheetName };
  }
}
