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
import {
  toWireValue,
  type Value,
  type ValueName,
  type VnToCvn,
} from "../01_SpreadsheetSchema/valueSchemas";
import { CellRaw } from "../02_SpreadsheetRaw/CellRaw";
import { CellBaseIdentified } from "./ClassBases/CellBaseIdentified";
import type { CellChange } from "./ClassTypes/StateIdentified";
import { ColumnIdentified } from "./ColumnIdentified";

export class CellIdentified<
  VN extends ValueName = ValueName,
> extends CellBaseIdentified<VN> {
  get column(): ColumnIdentified<VN> {
    return new ColumnIdentified(this.cellIdentifiedProps);
  }
  get raw(): CellRaw<VnToCvn<VN>> {
    return new CellRaw<VnToCvn<VN>>({
      ...this.cellIdentifiedProps,
      rowIndex: this.rowIndex,
      colIndex: this.column.colIndex,
    });
  }
  get isActive(): boolean {
    return this.raw.isActive;
  }
  prepFetch(): this {
    this.fetchTargets.push({
      kind: "singleCell",
      row: this.rowIndex,
      column: this.columnId,
    });
    return this;
  }
  // Identified is the lowest tier that knows the value name, so the blank is read here.
  valueOrEmpty(): Value<VN> {
    const value = this.raw.valueOrEmpty();
    const blankReadsAs = this.schema.valTrait("blankReadsAs");
    if (value === "" && blankReadsAs !== null) {
      return blankReadsAs as Value<VN>;
    }
    return value as Value<VN>;
  }
  valueNotEmpty(): NotEmpty<Value<VN>> {
    const value = this.valueOrEmpty();
    if (value === "") {
      throw new Error(
        `Value for column "${this.columnId}" in row ${this.rowIndex} is empty.`,
      );
    } else {
      return value as NotEmpty<Value<VN>>;
    }
  }
  // Both halves merge into the one queued change, so the field mask names both.
  update({ value, backgroundColor }: CellChange<VN>): this {
    if (value !== undefined) this.updateValue(value);
    if (backgroundColor !== undefined) {
      this.updateBackgroundColor(backgroundColor);
    }
    return this;
  }
  updateValue(value: Value<VN>): this {
    this.schema.validateDataNotFormula();
    this.raw.updateValue(toWireValue(value));
    return this;
  }
  updateFormula(formula: string): this {
    this.schema.validateIsFormula();
    this.raw.updateFormula(formula);
    return this;
  }
  updateBackgroundColor(backgroundColor: RgbColor): this {
    this.raw.updateBackgroundColor(backgroundColor);
    return this;
  }
  addConditionalFormatRule(declaration: ConditionalFormatDeclaration): this {
    this.raw.addConditionalFormatRule(declaration);
    return this;
  }
  removeConditionalFormatRules(): this {
    this.raw.removeConditionalFormatRules();
    return this;
  }
  removeConditionalFormatRule(rule: ConditionalFormatRule): this {
    this.raw.removeConditionalFormatRule(rule);
    return this;
  }
  addEditWarning(declaration: EditWarningDeclaration = {}): this {
    this.raw.addEditWarning(declaration);
    return this;
  }
  addEditLock(declaration: EditLockDeclaration = {}): this {
    this.raw.addEditLock(declaration);
    return this;
  }
  removeEditProtections(): this {
    this.raw.removeEditProtections();
    return this;
  }
  removeEditProtection(protection: EditProtection): this {
    this.raw.removeEditProtection(protection);
    return this;
  }
  anchoredA1(colIndex = this.column.colIndex): string {
    return this.schema.anchoredA1(colIndex, this.rowIndex);
  }
  updateToDefault(): this {
    if (!this.schema.isFormula) {
      this.updateValue(this.schema.makeDefaultDataValue() as Value<VN>);
    }
    return this;
  }
}
