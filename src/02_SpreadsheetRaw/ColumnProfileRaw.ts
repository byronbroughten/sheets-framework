import type { CellValue } from "../00_Source/CellValues/cellValues";
import type { FrameworkValueName } from "../00_Source/CellValues/frameworkValueSchemas";
import type { TableColumnType } from "../00_Source/RawSource/RawSource";
import { type PrimitiveValueName, Val } from "../utils/Val";
import { ColumnBaseRaw } from "./ClassBases/ColumnBaseRaw";
import type { ColumnStateRaw, SampledFactsRaw } from "./ClassTypes/StateRaw";
import { ColumnRaw } from "./ColumnRaw";
import { TableRaw } from "./TableRaw";

export class ColumnProfileRaw extends ColumnBaseRaw {
  private get table(): TableRaw {
    return new TableRaw(this.tableRawProps);
  }
  private get column(): ColumnRaw {
    return new ColumnRaw(this.columnRawProps);
  }
  get header(): string {
    return this.column.headCell("header").valueOrEmpty();
  }
  get isFormula(): boolean {
    return this._sampledFacts.isFormula;
  }
  get numberFormatType(): string | undefined {
    return this._sampledFacts.numberFormatType;
  }
  get dataValidationConditionType(): string | undefined {
    return this._sampledFacts.dataValidationConditionType;
  }
  get topValue(): CellValue {
    return this._sampledFacts.topValue;
  }
  get valueValidationStrings(): string[] {
    return this._tableColumnState()?.validationValues ?? [];
  }
  get validationConditionType(): string | undefined {
    return this._tableColumnState()?.validationConditionType;
  }
  get columnType(): string | undefined {
    return this._tableColumnState()?.columnType;
  }
  valueTitle(): string {
    return this.declaredValueTitle() ?? this._actualPrimitiveValueName();
  }
  declaredValueTitle(): string | undefined {
    if (this.header === this.schema.idHeader) {
      return "id";
    }
    return (
      this.validationValueTitle() ??
      this._columnTypeValueName() ??
      this._booleanValidationValueName() ??
      this._compatibleNumberFormatValueName()
    );
  }
  validationValueTitle(): string | undefined {
    for (const rawValue of this.valueValidationStrings) {
      const match = rawValue.match(/^=valueConfig\[(.+)\]$/);
      if (!match) continue;
      return Val.assert(match[1], "value title match");
    }
    return undefined;
  }
  private get _sampledFacts(): SampledFactsRaw {
    const facts = this.columnState?.sampledFacts;
    if (facts === undefined) {
      throw new Error(
        `No sampled facts for column index ${this.colIndex} of sheet ${this.sheetLabel}: ` +
          `nothing fetched its top data row in full. A column that was fetched and is ` +
          `simply empty reports blank facts instead.`,
      );
    }
    return facts;
  }
  // Table column properties are only trustworthy once the Table itself is known.
  private _tableColumnState(): ColumnStateRaw | undefined {
    this.table.assertTableIsKnown();
    return this.columnState;
  }
  private _columnTypeValueName(): FrameworkValueName | undefined {
    const { columnType } = this;
    if (columnType === undefined || !isNamedColumnType(columnType)) {
      return undefined;
    }
    return columnTypeValueNames[columnType];
  }
  private _booleanValidationValueName(): "checkbox" | undefined {
    if (this.validationConditionType === "BOOLEAN") {
      return "checkbox";
    }
    if (this.dataValidationConditionType === "BOOLEAN") {
      return "checkbox";
    }
    return undefined;
  }
  private _compatibleNumberFormatValueName(): PrimitiveValueName | undefined {
    const formatName = this._numberFormatValueName();
    if (formatName === undefined) {
      return undefined;
    }
    return this._actualPrimitiveValueName() === formatName
      ? formatName
      : undefined;
  }
  private _actualPrimitiveValueName(): PrimitiveValueName {
    const value = this.topValue;
    if (typeof value === "boolean") {
      return "boolean";
    }
    if (typeof value === "number") {
      return this._numberFormatValueName() === "date" ? "date" : "number";
    }
    if (value === "") {
      return this._numberFormatValueName() ?? "string";
    }
    return "string";
  }
  private _numberFormatValueName(): PrimitiveValueName | undefined {
    const formatType = this.numberFormatType;
    if (
      formatType === undefined ||
      formatType === "NUMBER_FORMAT_TYPE_UNSPECIFIED"
    ) {
      return undefined;
    }
    return numberFormatValueNames[formatType];
  }
}

// DROPDOWN and COLUMN_TYPE_UNSPECIFIED are absent: neither says what a column holds.
const columnTypeValueNames: Partial<
  Record<TableColumnType, FrameworkValueName>
> = {
  DOUBLE: "number",
  CURRENCY: "number",
  PERCENT: "number",
  DATE: "date",
  TIME: "number",
  DATE_TIME: "number",
  TEXT: "string",
  FILES_CHIP: "string",
  PEOPLE_CHIP: "string",
  FINANCE_CHIP: "string",
  PLACE_CHIP: "string",
  RATINGS_CHIP: "string",
  // Declaring the type is what earns the never-blank guarantee; a sampled boolean doesn't.
  BOOLEAN: "checkbox",
};

function isNamedColumnType(
  columnType: string,
): columnType is keyof typeof columnTypeValueNames {
  return Object.hasOwn(columnTypeValueNames, columnType);
}

const numberFormatValueNames: Record<string, PrimitiveValueName> = {
  DATE: "date",
  TIME: "number",
  DATE_TIME: "number",
  NUMBER: "number",
  CURRENCY: "number",
  PERCENT: "number",
  SCIENTIFIC: "number",
  TEXT: "string",
};
