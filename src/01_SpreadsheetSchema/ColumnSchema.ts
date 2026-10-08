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
import { SpreadsheetBaseSchema } from "./SpreadsheetBaseSchema";
import type { TableName } from "./tableConfigsTypes";
import { TableSchema, type TableSchemaProps } from "./TableSchema";
import { getValTrait, type ValueSchema } from "./valueSchemas";

interface ColumnSchemaProps<
  TN extends TableName,
  CN extends ColumnName<TN>,
> extends TableSchemaProps<TN> {
  columnId: string;
  columnName: CN;
}

export class ColumnSchema<
  TN extends TableName = TableName,
  CN extends ColumnName<TN> = ColumnName<TN>,
> extends SpreadsheetBaseSchema {
  readonly sheetGid: number;
  readonly tableName: TN;
  readonly columnId: string;
  readonly columnName: CN;
  constructor({
    sheetGid,
    tableName,
    columnId,
    columnName,
  }: ColumnSchemaProps<TN, CN>) {
    super();
    this.sheetGid = sheetGid;
    this.tableName = tableName;
    this.columnId = columnId;
    this.columnName = columnName;
  }
  static fromColumnName<TN extends TableName, CN extends ColumnName<TN>>(
    tableName: TN,
    columnName: CN,
  ): ColumnSchema<TN, CN> {
    return TableSchema.fromSheetName(tableName).columnByName(columnName);
  }
  static fromColumnId(tableName: TableName, columnId: string): ColumnSchema {
    return TableSchema.fromSheetName(tableName).columnById(columnId);
  }
  get sheet(): TableSchema<TN> {
    return new TableSchema({
      sheetGid: this.sheetGid,
      tableName: this.tableName,
    });
  }
  trait<TK extends keyof ColumnConfig>(
    key: TK,
  ): ColumnConfigAt<TN, CN>[TK & keyof ColumnConfigAt<TN, CN>] {
    return getColumnTraitById(
      this.tableName,
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
      this.tableName,
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
