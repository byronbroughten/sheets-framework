import type { CellValue, NotEmpty } from "../00_Source/CellValues/cellValues";
import { type Value } from "../01_SpreadsheetSchema/valueSchemas";
import { RowRaw } from "../02_SpreadsheetRaw/RowRaw";
import { CellIdentified } from "./CellIdentified";
import type { RowIdentifiedProps } from "./ClassBases/RowBaseIdentified";
import { RowCommonIdentified } from "./ClassBases/RowCommonIdentified";
import { SheetIdentified } from "./SheetIdentified";

export class RowIdentified extends RowCommonIdentified {
  constructor(props: RowIdentifiedProps) {
    super(props);
    void this.raw;
  }
  get sheet(): SheetIdentified {
    return new SheetIdentified(this.sheetIdentifiedProps);
  }
  get raw(): RowRaw {
    return new RowRaw(this.rowIdentifiedProps);
  }
  get activeValueArr(): CellValue[] {
    return this.raw.activeValueArr;
  }
  valueOrEmpty(columnId: string): Value {
    return this.cell(columnId).valueOrEmpty();
  }
  valueNotEmpty(columnId: string): NotEmpty<Value> {
    return this.cell(columnId).valueNotEmpty();
  }
  updateValue(columnId: string, value: Value): this {
    this.cell(columnId).updateValue(value);
    return this;
  }
  cell(columnId: string): CellIdentified {
    return new CellIdentified({
      ...this.rowIdentifiedProps,
      columnId,
    });
  }
  updateToDefault(...columnIds: string[]): RowIdentified {
    columnIds.forEach((columnId) => this.cell(columnId).updateToDefault());
    return this;
  }
  get activeColumnIds(): string[] {
    return [...this.raw.rowState.keys()].map((colIndex) =>
      this.sheet.meta.columnIdByIndex(colIndex),
    );
  }
  get isActive(): boolean {
    return this.raw.rowIsActive();
  }
  get isQueuedForDelete(): boolean {
    return this.raw.isQueuedForDelete;
  }
  // Raw decides, since a checkbox column's blank reads as false and an unread row isn't empty.
  get isBlank(): boolean {
    if (!this.isActive) return false;
    return this._nonFormulaCellsActive.every((cell) => cell.raw.isEmpty);
  }
  get isReusable(): boolean {
    return this.isBlank && !this.raw.isReserved;
  }
  reserve(): void {
    this.raw.reserve();
  }
  // A blank row needs no writes, but its reservation must lift either way.
  clearValues(): this {
    if (!this.isBlank) {
      this.sheet.nonFormulaColumnIds.forEach((columnId) => {
        this.updateValue(columnId, "");
      });
    }
    this.raw.release();
    return this;
  }
  delete(): void {
    if (this.sheet.raw.isDownToLastDataRow) {
      this.clearValues();
    } else {
      this.raw.delete();
    }
  }
  private get _nonFormulaCellsActive(): CellIdentified[] {
    return this.sheet.nonFormulaColumnIds
      .map((columnId) => this.cell(columnId))
      .filter((cell) => cell.isActive);
  }
}
