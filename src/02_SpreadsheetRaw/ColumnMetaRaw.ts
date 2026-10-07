import type {
  CellValue,
  CellValueName,
} from "../00_Source/CellValues/cellValues";
import type { FrameworkValueName } from "../00_Source/CellValues/frameworkValueSchemas";
import type {
  GridCellSnapshot,
  TableColumnType,
} from "../00_Source/RawSource/RawSource";
import { type PrimitiveValueName, Val } from "../utils/Val";
import { ColumnBaseRaw } from "./ClassBases/ColumnBaseRaw";
import type {
  ActiveFactsRaw,
  ColumnStateRaw,
  TableEndColumnHeadCells,
} from "./ClassTypes/StateRaw";
import { ColumnRaw } from "./ColumnRaw";
import { SheetMetaRaw } from "./SheetMetaRaw";

export class ColumnMetaRaw<
  VN extends CellValueName = CellValueName,
> extends ColumnBaseRaw {
  get table(): SheetMetaRaw {
    return new SheetMetaRaw(this.tableRawProps);
  }
  get primary(): ColumnRaw<VN> {
    return new ColumnRaw<VN>(this.columnRawProps);
  }
  get activeHeader(): string {
    return this.primary.headCell("header").valueOrEmpty();
  }
  get activeIsFormula(): boolean {
    return this._activeFacts.isFormula;
  }
  get activeNumberFormatType(): string | undefined {
    return this._activeFacts.numberFormatType;
  }
  get activeDataValidationConditionType(): string | undefined {
    return this._activeFacts.dataValidationConditionType;
  }
  get activeTopValue(): CellValue {
    return this._activeFacts.topValue;
  }
  private get _activeFacts(): ActiveFactsRaw {
    const facts = this.columnState?.activeFacts;
    if (facts === undefined) {
      throw new Error(
        `No active facts for column index ${this.colIndex} of sheet ${this.sheetLabel}: ` +
          `nothing fetched its top data row in full. A column that was fetched and is ` +
          `simply empty reports blank facts instead.`,
      );
    }
    return facts;
  }
  get valueValidationStrings(): string[] {
    return this._tableColumnState()?.validationValues ?? [];
  }
  get validationConditionType(): string | undefined {
    return this._tableColumnState()?.validationConditionType;
  }
  get activeColumnType(): string | undefined {
    return this._tableColumnState()?.columnType;
  }
  updateColumnType(columnType: TableColumnType): this {
    this.table.assertTableIsKnown();
    this.table.queueTableWrite({
      action: "updateColumnType",
      colIndex: this.colIndex,
      columnType,
    });
    this._ensureColumnState(this.colIndex).columnType = columnType;
    return this;
  }
  // Table column properties are only trustworthy once the Table itself is known.
  private _tableColumnState(): ColumnStateRaw | undefined {
    this.table.assertTableIsKnown();
    return this.columnState;
  }
  initHeadCells({
    columnId,
    header,
    groupHeading1,
  }: TableEndColumnHeadCells): this {
    const { primary } = this;
    primary.headCell("columnId").updateValue(columnId);
    primary.headCell("header").updateValue(header);
    if (groupHeading1 !== undefined) {
      primary.headCell("groupHeading1").updateValue(groupHeading1);
    }
    return this;
  }
  integrateActiveFacts(cell: GridCellSnapshot | undefined): void {
    this._ensureColumnState(this.colIndex).activeFacts = {
      isFormula: cell?.isFormula ?? false,
      numberFormatType: cell?.numberFormatType,
      dataValidationConditionType: cell?.dataValidationConditionType,
      topValue: cell?.value ?? "", // from the payload, so a deleted top data row still describes the column
    };
  }
  // Gap-filling only, so a fact the payload described always wins.
  ensureActiveFacts(): void {
    if (this.columnState?.activeFacts !== undefined) return;
    if (!this.primary.topCell.inWorking) return; // no top data row to sample
    this.integrateActiveFacts(undefined);
  }
  activeValueTitle(): string {
    return this.activeDeclaredValueTitle() ?? this._actualPrimitiveValueName();
  }
  activeDeclaredValueTitle(): string | undefined {
    if (this.activeHeader === this.schema.idHeader) {
      return "id";
    }
    return (
      this.activeValidationValueTitle() ??
      this._columnTypeValueName() ??
      this._booleanValidationValueName() ??
      this._compatibleNumberFormatValueName()
    );
  }
  activeValidationValueTitle(): string | undefined {
    for (const rawValue of this.valueValidationStrings) {
      const match = rawValue.match(/^=valueConfig\[(.+)\]$/);
      if (!match) continue;
      return Val.assert(match[1], "value title match");
    }
    return undefined;
  }
  private _columnTypeValueName(): FrameworkValueName | undefined {
    const columnType = this.activeColumnType;
    if (columnType === undefined || !isNamedColumnType(columnType)) {
      return undefined;
    }
    return columnTypeValueNames[columnType];
  }
  private _booleanValidationValueName(): "checkbox" | undefined {
    if (this.validationConditionType === "BOOLEAN") {
      return "checkbox";
    }
    if (this.activeDataValidationConditionType === "BOOLEAN") {
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
    const value = this.activeTopValue;
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
    const formatType = this.activeNumberFormatType;
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
