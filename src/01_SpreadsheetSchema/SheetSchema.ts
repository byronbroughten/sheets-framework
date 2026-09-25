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
  type SheetName,
} from "./sheetConfigsTypes";
import { SpreadsheetBaseSchema } from "./SpreadsheetBaseSchema";

function sheetNameFromGid(sheetGid: number): SheetName {
  const byGid = sheetConfigsByGid();
  if (!byGid.has(sheetGid)) {
    throw new Error(
      `Invalid sheetGid: ${sheetGid}. Must be one of: ${[...byGid.keys()].join(", ")}`,
    );
  }
  return getSheetTraitByGid(sheetGid, "sheetName") as SheetName;
}

export interface SheetSchemaProps<SN extends SheetName> {
  sheetGid: number;
  sheetName: SN;
}

export class SheetSchema<
  SN extends SheetName = SheetName,
> extends SpreadsheetBaseSchema {
  readonly sheetGid: number;
  readonly sheetName: SN;
  constructor({ sheetGid, sheetName }: SheetSchemaProps<SN>) {
    super();
    this.sheetGid = sheetGid;
    this.sheetName = sheetName;
  }
  static fromSheetName<SN extends SheetName>(sheetName: SN): SheetSchema<SN> {
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
  get columnNames(): ColumnName<SN>[] {
    return getSheetColumnNames(this.sheetName);
  }
  get nonFormulaColumnIds(): string[] {
    return [...this.columnIds].filter((columnId) => {
      return !getColumnTraitById(this.sheetGid, columnId, "isFormula");
    });
  }
  // Goes by gid, the only O(1) columnId -> columnName index; by name would scan the sheet.
  colNameByColumnId(columnId: string): ColumnName<SN> {
    return getColumnTraitById(
      this.sheetGid,
      columnId,
      "columnName",
    ) as ColumnName<SN>;
  }
  columnByName<CN extends ColumnName<SN>>(
    columnName: CN,
  ): ColumnSchema<SN, CN> {
    return new ColumnSchema({
      ...this.sheetSchemaProps,
      columnName,
      columnId: getColumnTraitByName(this.sheetName, columnName, "columnId"),
    });
  }
  columnById(columnId: string): ColumnSchema<SN, ColumnName<SN>> {
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
    columnSpecifier: ColumnName<SN> | ColumnName<SN>[] | "allColumns",
  ): ColumnName<SN>[] {
    if (columnSpecifier === "allColumns") {
      return this.columnNames;
    } else if (Array.isArray(columnSpecifier)) {
      return columnSpecifier;
    } else {
      return [columnSpecifier];
    }
  }
  private get sheetSchemaProps(): SheetSchemaProps<SN> {
    return { sheetGid: this.sheetGid, sheetName: this.sheetName };
  }
}
