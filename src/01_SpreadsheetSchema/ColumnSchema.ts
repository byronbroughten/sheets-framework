import type { ValueSchemaKey } from "../00_Source/CellValues/valueSchema";
import {
  type ColumnConfig,
  type ColumnConfigAt,
  type ColumnFullName,
  type ColumnName,
  type ColumnValue,
  getColumnTraitById,
  type MakeColumnFullName,
} from "./columnConfigsTypes";
import type { TableName } from "./sheetConfigsTypes";
import { SheetSchema, type SheetSchemaProps } from "./SheetSchema";
import { SpreadsheetBaseSchema } from "./SpreadsheetBaseSchema";
import { getValTrait, type ValueSchema } from "./valueSchemas";

interface ColumnSchemaProps<
  TN extends TableName,
  CN extends ColumnName<TN>,
> extends SheetSchemaProps<TN> {
  columnId: string;
  columnName: CN;
}

export class ColumnSchema<
  TN extends TableName = TableName,
  CN extends ColumnName<TN> = ColumnName<TN>,
> extends SpreadsheetBaseSchema {
  readonly sheetGid: number;
  readonly sheetName: TN;
  readonly columnId: string;
  readonly columnName: CN;
  constructor({
    sheetGid,
    sheetName,
    columnId,
    columnName,
  }: ColumnSchemaProps<TN, CN>) {
    super();
    this.sheetGid = sheetGid;
    this.sheetName = sheetName;
    this.columnId = columnId;
    this.columnName = columnName;
  }
  static fromColumnName<TN extends TableName, CN extends ColumnName<TN>>(
    sheetName: TN,
    columnName: CN,
  ): ColumnSchema<TN, CN> {
    return SheetSchema.fromSheetName(sheetName).columnByName(columnName);
  }
  static fromColumnId(sheetGid: number, columnId: string): ColumnSchema {
    return SheetSchema.fromSheetGid(sheetGid).columnById(columnId);
  }
  get sheet(): SheetSchema<TN> {
    return new SheetSchema({
      sheetGid: this.sheetGid,
      sheetName: this.sheetName,
    });
  }
  trait<TK extends keyof ColumnConfig>(
    key: TK,
  ): ColumnConfigAt<TN, CN>[TK & keyof ColumnConfigAt<TN, CN>] {
    return getColumnTraitById(
      this.sheetGid,
      this.columnId,
      key,
    ) as ColumnConfigAt<TN, CN>[TK & keyof ColumnConfigAt<TN, CN>];
  }
  get valueName(): ColumnConfigAt<TN, CN>["valueName"] {
    return this.trait("valueName");
  }
  valTrait<VK extends ValueSchemaKey>(
    key: VK,
  ): ValueSchema<ColumnConfigAt<TN, CN>["valueName"]>[VK] {
    return getValTrait(this.valueName, key);
  }
  get isFormula(): boolean {
    return this.trait("isFormula");
  }
  get emptyValueAllowed(): boolean {
    return this.trait("emptyValueAllowed");
  }
  get fullName(): MakeColumnFullName<TN, CN> & ColumnFullName {
    return this.combineNames(
      this.sheetName,
      this.columnName as string,
    ) as MakeColumnFullName<TN, CN> & ColumnFullName;
  }
  makeRowId(): string {
    return this.sheet.makeRowId();
  }
  makeDefaultDataValue(): ColumnValue<TN, CN> {
    if ((this.columnName as string) === "id") {
      return this.makeRowId() as ColumnValue<TN, CN>;
    } else {
      return this.valTrait("makeDefault")() as ColumnValue<TN, CN>;
    }
  }
  validate(value: unknown): ColumnValue<TN, CN> | "" {
    if (this.emptyValueAllowed && value === "") {
      return value;
    } else {
      return this.valTrait("strictValidate")(value);
    }
  }
  validateDataNotFormula(): void {
    if (this.isFormula) {
      throw new Error(
        `Column with id "${this.columnId}" is a formula column and cannot be used for this operation.`,
      );
    }
  }
  validateIsFormula(): void {
    if (this.isFormula) return;
    throw new Error(
      `Column with id "${this.columnId}" is not a formula column and cannot be used for this operation.`,
    );
  }
}
