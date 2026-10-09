import type { NotEmpty } from "../00_Source/CellValues/cellValues";
import type {
  ConditionalFormatDeclaration,
  ConditionalFormatRule,
} from "../00_Source/RawSource/ConditionalFormat";
import type {
  EditLockDeclaration,
  EditProtection,
  EditWarningDeclaration,
} from "../00_Source/RawSource/EditProtection";
import type { RgbColor } from "../00_Source/RawSource/RgbColor";
import type {
  ColumnIsFormula,
  ColumnName,
  ColumnValue,
  ColumnValueDeclared,
  ColumnValueName,
} from "../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import type { TableName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import type {
  Value,
  ValueName,
  VnToCvn,
} from "../01_SpreadsheetSchema/configReaders/valueSchemas";
import type { CellRaw } from "../02_SpreadsheetRaw/CellRaw";
import type { CellIdentified } from "../03_SpreadsheetIdentified/CellIdentified";
import { CellBaseNamed } from "./ClassBases/CellBaseNamed";
import { ColumnNamed } from "./ColumnNamed";

export class CellNamed<
  TN extends TableName,
  CN extends ColumnName<TN> = ColumnName<TN>,
> extends CellBaseNamed<TN, CN> {
  get column(): ColumnNamed<TN, CN> {
    return new ColumnNamed(this.columnNamedProps);
  }
  get identified(): CellIdentified<ColumnValueName<TN, CN>> {
    return this.column.identified.cell(this.rowIndex);
  }
  get raw(): CellRaw<VnToCvn<ColumnValueName<TN, CN>>> {
    return this.identified.raw;
  }
  get inWorking(): boolean {
    return this.identified.inWorking;
  }
  valueOrEmpty(): ColumnValue<TN, CN> {
    return this.identified.valueOrEmpty();
  }
  // Checked here, not delegated, so the message names the column the caller wrote.
  valueNotEmpty(): NotEmpty<ColumnValue<TN, CN>> {
    const value = this.valueOrEmpty();
    if (value === "") {
      throw new Error(
        `Column "${this.columnName}" of sheet "${this.tableName}" is empty in ${this.identified.raw.rowLabel(this.rowIndex)}.`,
      );
    } else {
      return value as NotEmpty<ColumnValue<TN, CN>>;
    }
  }
  value(): ColumnValueDeclared<TN, CN> {
    if (this.schema.emptyValueAllowed) {
      return this.valueOrEmpty() as ColumnValueDeclared<TN, CN>;
    } else {
      return this.valueNotEmpty() as ColumnValueDeclared<TN, CN>;
    }
  }
  updateValue(value: ColumnValue<TN, CN>): this {
    this.identified.updateValue(value);
    return this;
  }
  updateFormula(
    formula: ColumnIsFormula<TN, CN> extends true ? string : never,
  ): this {
    this.identified.updateFormula(formula);
    return this;
  }
  updateBackgroundColor(backgroundColor: RgbColor): this {
    this.identified.updateBackgroundColor(backgroundColor);
    return this;
  }
  updateToDefault(): this {
    this.identified.updateToDefault();
    return this;
  }
  setValueType(valueName: ValueName, value: Value): this {
    if (this.schema.valueName !== valueName) {
      throw new Error(
        `Value name ${valueName} does not match varb value name ${this.schema.valueName}`,
      );
    }
    const validated = this.schema.validate(value);
    this.updateValue(validated as ColumnValue<TN, CN>);
    return this;
  }
  addConditionalFormatRule(declaration: ConditionalFormatDeclaration): this {
    this.identified.addConditionalFormatRule(declaration);
    return this;
  }
  removeConditionalFormatRules(): this {
    this.identified.removeConditionalFormatRules();
    return this;
  }
  removeConditionalFormatRule(rule: ConditionalFormatRule): this {
    this.identified.removeConditionalFormatRule(rule);
    return this;
  }
  addEditWarning(declaration: EditWarningDeclaration = {}): this {
    this.identified.addEditWarning(declaration);
    return this;
  }
  addEditLock(declaration: EditLockDeclaration = {}): this {
    this.identified.addEditLock(declaration);
    return this;
  }
  removeEditProtections(): this {
    this.identified.removeEditProtections();
    return this;
  }
  removeEditProtection(protection: EditProtection): this {
    this.identified.removeEditProtection(protection);
    return this;
  }
  anchoredA1(columnName: ColumnName<TN> = this.columnName): string {
    const colIndex = this.column.table.column(columnName).identified.colIndex;
    return this.identified.anchoredA1(colIndex);
  }
}
