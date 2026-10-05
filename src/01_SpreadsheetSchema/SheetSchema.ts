import {
  type ColumnName,
  getColumnTraitById,
  getColumnTraitByName,
  getSheetColumnIds,
  getSheetColumnNames,
} from "./columnConfigsTypes";
import { ColumnSchema } from "./ColumnSchema";
import { dimensionIds } from "./dimensionIds";
import {
  getSheetTraitByGid,
  getSheetTraitByName,
  type SheetConfig,
  sheetConfigsByGid,
  type TableName,
} from "./sheetConfigsTypes";
import { SpreadsheetBaseSchema } from "./SpreadsheetBaseSchema";

function sheetNameFromGid(sheetGid: number): TableName {
  const byGid = sheetConfigsByGid();
  if (!byGid.has(sheetGid)) {
    throw new Error(
      `Invalid sheetGid: ${sheetGid}. Must be one of: ${[...byGid.keys()].join(", ")}`,
    );
  }
  return getSheetTraitByGid(sheetGid, "sheetName") as TableName;
}

export interface SheetSchemaProps<TN extends TableName> {
  sheetGid: number;
  sheetName: TN;
}

export class SheetSchema<
  TN extends TableName = TableName,
> extends SpreadsheetBaseSchema {
  readonly sheetGid: number;
  readonly sheetName: TN;
  constructor({ sheetGid, sheetName }: SheetSchemaProps<TN>) {
    super();
    this.sheetGid = sheetGid;
    this.sheetName = sheetName;
  }
  static fromSheetName<TN extends TableName>(sheetName: TN): SheetSchema<TN> {
    return new SheetSchema({
      sheetName,
      sheetGid: getSheetTraitByName(sheetName, "sheetGid"),
    });
  }
  static fromSheetGid(sheetGid: number): SheetSchema {
    return new SheetSchema({
      sheetGid,
      sheetName: sheetNameFromGid(sheetGid),
    });
  }
  trait<TK extends keyof SheetConfig>(key: TK): SheetConfig[TK] {
    return getSheetTraitByGid(this.sheetGid, key);
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
  private get sheetSchemaProps(): SheetSchemaProps<TN> {
    return { sheetGid: this.sheetGid, sheetName: this.sheetName };
  }
}
