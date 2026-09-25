import type {
  CellValue,
  CellValueName,
  UniformRowName,
  UniformRowValue,
  UniformRowValueName,
} from "../00_Source/CellValues/cellValues";
import type { FrameworkValueName } from "../00_Source/CellValues/frameworkValueSchemas";
import type {
  GridCellSnapshot,
  TableColumnType,
} from "../00_Source/RawSource/RawSource";
import { dimensionIds } from "../01_SpreadsheetSchema/dimensionIds";
import { type PrimitiveValueName, Val } from "../utils/Val";
import { CellRaw } from "./CellRaw";
import { ColumnBaseRaw } from "./ClassBases/ColumnBaseRaw";
import type { ActiveFactsRaw, ColumnStateRaw } from "./ClassTypes/StateRaw";
import { ColumnRaw } from "./ColumnRaw";
import { SheetMetaRaw } from "./SheetMetaRaw";

export class ColumnMetaRaw<
  VN extends CellValueName = CellValueName,
> extends ColumnBaseRaw {
  get sheet(): SheetMetaRaw {
    return new SheetMetaRaw(this.sheetRawProps);
  }
  get primary(): ColumnRaw<VN> {
    return new ColumnRaw<VN>(this.columnRawProps);
  }
  get activeHeader(): string {
    return this.uniformCell("tableHeader").valueOrEmpty();
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
    this.sheet.activeTable.validateColIndexNotStale(this.colIndex);
    this.updateRequests.updateTableColumnType.push({
      kind: "updateTableColumnType",
      sheetId: this.sheetGid,
      tableId: this.sheet.activeTable.tableId,
      // Google's Table columnIndex is table-relative; colIndex is sheet-absolute.
      columnIndex: this.colIndex - this.sheet.activeTable.startColumnIndex,
      columnType,
    });
    this._ensureColumnState(this.colIndex).columnType = columnType;
    return this;
  }
  // Table column properties are only trustworthy once the Table itself is known.
  private _tableColumnState(): ColumnStateRaw | undefined {
    this.sheet.activeTable.assertKnown();
    return this.columnState;
  }
  uniformCell<UN extends UniformRowName>(
    rowName: UN,
  ): CellRaw<UniformRowValueName<UN>> {
    return new CellRaw<UniformRowValueName<UN>>({
      ...this.columnRawProps,
      rowIndex: this.schema.uniformRowIndex(rowName),
    });
  }
  initUniformCells({
    idPrefix,
    header,
  }: {
    idPrefix: string;
    header: string;
  }): this {
    const columnId = dimensionIds.col(idPrefix);
    this.uniformCell("columnId").updateValue(columnId);
    this.uniformCell("tableHeader").updateValue(header);
    return this;
  }
  updateUniformCell<UN extends UniformRowName>(
    rowName: UN,
    newValue: UniformRowValue<UN>,
  ): this {
    this.uniformCell(rowName).updateValue(newValue);
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
    if (!this.primary.topCell.isActive) return; // no top data row to sample
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
