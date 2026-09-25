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
import type { SheetName } from "./sheetConfigsTypes";
import { SheetSchema, type SheetSchemaProps } from "./SheetSchema";
import { SpreadsheetBaseSchema } from "./SpreadsheetBaseSchema";
import { getValTrait, type ValueSchema } from "./valueSchemas";

interface ColumnSchemaProps<
  SN extends SheetName,
  CN extends ColumnName<SN>,
> extends SheetSchemaProps<SN> {
  columnId: string;
  columnName: CN;
}

export class ColumnSchema<
  SN extends SheetName = SheetName,
  CN extends ColumnName<SN> = ColumnName<SN>,
> extends SpreadsheetBaseSchema {
  readonly sheetGid: number;
  readonly sheetName: SN;
  readonly columnId: string;
  readonly columnName: CN;
  constructor({
    sheetGid,
    sheetName,
    columnId,
    columnName,
  }: ColumnSchemaProps<SN, CN>) {
    super();
    this.sheetGid = sheetGid;
    this.sheetName = sheetName;
    this.columnId = columnId;
    this.columnName = columnName;
  }
  static fromColumnName<SN extends SheetName, CN extends ColumnName<SN>>(
    sheetName: SN,
    columnName: CN,
  ): ColumnSchema<SN, CN> {
    return SheetSchema.fromSheetName(sheetName).columnByName(columnName);
  }
  static fromColumnId(sheetGid: number, columnId: string): ColumnSchema {
    return SheetSchema.fromSheetGid(sheetGid).columnById(columnId);
  }
  get sheet(): SheetSchema<SN> {
    return new SheetSchema({
      sheetGid: this.sheetGid,
      sheetName: this.sheetName,
    });
  }
  trait<TK extends keyof ColumnConfig>(
    key: TK,
  ): ColumnConfigAt<SN, CN>[TK & keyof ColumnConfigAt<SN, CN>] {
    return getColumnTraitById(
      this.sheetGid,
      this.columnId,
      key,
    ) as ColumnConfigAt<SN, CN>[TK & keyof ColumnConfigAt<SN, CN>];
  }
  get valueName(): ColumnConfigAt<SN, CN>["valueName"] {
    return this.trait("valueName");
  }
  valTrait<VK extends ValueSchemaKey>(
    key: VK,
  ): ValueSchema<ColumnConfigAt<SN, CN>["valueName"]>[VK] {
    return getValTrait(this.valueName, key);
  }
  get isFormula(): boolean {
    return this.trait("isFormula");
  }
  get emptyValueAllowed(): boolean {
    return this.trait("emptyValueAllowed");
  }
  get fullName(): MakeColumnFullName<SN, CN> & ColumnFullName {
    return this.combineNames(
      this.sheetName,
      this.columnName as string,
    ) as MakeColumnFullName<SN, CN> & ColumnFullName;
  }
  makeRowId(): string {
    return this.sheet.makeRowId();
  }
  makeDefaultDataValue(): ColumnValue<SN, CN> {
    if ((this.columnName as string) === "id") {
      return this.makeRowId() as ColumnValue<SN, CN>;
    } else {
      return this.valTrait("makeDefault")() as ColumnValue<SN, CN>;
    }
  }
  validate(value: unknown): ColumnValue<SN, CN> | "" {
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
