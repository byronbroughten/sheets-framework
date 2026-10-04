import type {
  CellValue,
  CellValueName,
} from "../00_Source/CellValues/cellValues";
import type {
  ConditionalFormatDeclaration,
  ConditionalFormatRule,
} from "../00_Source/RawSource/ConditionalFormat";
import type {
  EditLockDeclaration,
  EditProtection,
  EditWarningDeclaration,
} from "../00_Source/RawSource/EditProtection";
import type {
  BoundedGridRange,
  GridRangeProps,
} from "../00_Source/RawSource/RawSource";
import {
  SheetIndex,
  type SheetRowIndex,
} from "../00_Source/RawSource/SheetIndex";
import { Arr } from "../utils/Arr";
import { CellRaw, validateFormulaString } from "./CellRaw";
import { ColumnBaseRaw } from "./ClassBases/ColumnBaseRaw";
import type { CellFill, FindReplaceTerms } from "./ClassTypes/StateRaw";
import { ColumnMetaRaw } from "./ColumnMetaRaw";
import { SheetRaw } from "./SheetRaw";
import { SpreadsheetRaw } from "./SpreadsheetRaw";

export class ColumnRaw<
  VN extends CellValueName = CellValueName,
> extends ColumnBaseRaw {
  get ss(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
  get sheet(): SheetRaw {
    return new SheetRaw(this.sheetRawProps);
  }
  get meta(): ColumnMetaRaw<VN> {
    return new ColumnMetaRaw<VN>(this.columnRawProps);
  }
  get valueArrOrEmpty(): (CellValue<VN> | "")[] {
    return this.sheet.rowIndexesActive.map((rowIndex) =>
      this.valueOrEmpty(rowIndex),
    );
  }
  get valueArrFilterEmpty(): CellValue<VN>[] {
    return this.valueArrOrEmpty.filter(
      (value): value is CellValue<VN> => value !== "",
    );
  }
  get topCell(): CellRaw<VN> {
    return this.cell(0);
  }
  dataGridRange(): BoundedGridRange {
    const { origin, dataRowCount } = this.sheet.activeTable;
    return {
      sheetId: this.sheetGid,
      startRowIndex: origin.sheetRowIndex(0),
      endRowIndex: origin.sheetRowIndex(dataRowCount),
      startColumnIndex: origin.sheetColIndex(this.colIndex),
      endColumnIndex: origin.sheetColIndex(this.colIndex + 1),
    };
  }
  gridRangeFromRow(startRowIndex: number): GridRangeProps {
    return this._gridRangeFromSheetRow(
      this.tableOrigin().sheetRowIndex(startRowIndex),
    );
  }
  get cellIndexesActive(): number[] {
    return this.sheet.rowIndexesActive;
  }
  get cellIndexesFull(): number[] {
    return this.sheet.rowIndexesFull;
  }
  cell(rowIndex: number): CellRaw<VN> {
    return new CellRaw<VN>({
      ...this.columnRawProps,
      rowIndex,
    });
  }
  valueOrEmpty(rowIndex: number): CellValue<VN> | "" {
    return this.cell(rowIndex).valueOrEmpty();
  }
  updateValue(rowIndex: number, newValue: CellValue<VN>): this {
    this.cell(rowIndex).updateValue(newValue);
    return this;
  }
  // State is still mirrored row by row; only the queued request collapses.
  updateAllCells(change: Omit<CellFill<VN>, "formula">): this {
    this.sheet.activeTable.assertRowIndexesNotStale();
    this.sheet.validateNotPrunedToSelection();
    const { dataRowCount } = this.sheet.activeTable;
    const { value } = change;
    this.sheet.rowIndexesFull.forEach((rowIndex) => {
      const row = this.sheet.row(rowIndex);
      row.validateIsWritable();
      if (value !== undefined && row.rowIsActive()) {
        this.cell(rowIndex).setValueState(value);
      }
    });
    this.sheet.queueSheetWrite({
      action: "fillColumn",
      colIndex: this.colIndex,
      startRowIndex: 0,
      endRowIndex: dataRowCount,
      ...change,
    });
    return this;
  }
  updateAllFormulas(formula: string): this {
    this.sheet.activeTable.assertRowIndexesNotStale();
    validateFormulaString(formula);
    this.sheet.validateNotPrunedToSelection();
    const { dataRowCount } = this.sheet.activeTable;
    this.sheet.rowIndexesFull.forEach((rowIndex) => {
      this.sheet.row(rowIndex).validateIsWritable();
    });
    this.sheet.queueSheetWrite({
      action: "fillColumn",
      colIndex: this.colIndex,
      startRowIndex: 0,
      endRowIndex: dataRowCount,
      formula,
    });
    return this;
  }
  updateActiveCells(change: Omit<CellFill<VN>, "formula">): this {
    this.sheet.activeTable.assertRowIndexesNotStale();
    const rowIndexes = this.cellIndexesActive;
    const { value } = change;
    if (value !== undefined) {
      rowIndexes.forEach((rowIndex) => {
        this.cell(rowIndex).setValueState(value);
      });
    }
    Arr.contiguousRanges(rowIndexes).forEach(({ startIndex, endIndex }) => {
      this.sheet.queueSheetWrite({
        action: "fillColumn",
        colIndex: this.colIndex,
        startRowIndex: startIndex,
        endRowIndex: endIndex,
        ...change,
      });
    });
    return this;
  }
  updateActiveFormulas(formula: string): this {
    this.sheet.activeTable.assertRowIndexesNotStale();
    validateFormulaString(formula);
    Arr.contiguousRanges(this.cellIndexesActive).forEach(
      ({ startIndex, endIndex }) => {
        this.sheet.queueSheetWrite({
          action: "fillColumn",
          colIndex: this.colIndex,
          startRowIndex: startIndex,
          endRowIndex: endIndex,
          formula,
        });
      },
    );
    return this;
  }
  // Reaches every data row like a whole-column fill, so it takes the same guards.
  findReplace(terms: FindReplaceTerms): this {
    this.sheet.validateNotPrunedToSelection();
    this.ss.findReplace({ ...terms, scope: { range: this.dataGridRange() } });
    return this;
  }
  addConditionalFormatRule(declaration: ConditionalFormatDeclaration): this {
    this.sheet.addConditionalFormatRuleAt(this.dataGridRange(), declaration);
    return this;
  }
  removeConditionalFormatRules(): this {
    this.sheet.removeConditionalFormatRulesAt(this.dataGridRange());
    return this;
  }
  removeConditionalFormatRule(rule: ConditionalFormatRule): this {
    this.sheet.removeConditionalFormatRule(rule);
    return this;
  }
  addEditWarning(declaration: EditWarningDeclaration = {}): this {
    this.sheet.addEditWarningAt(this.dataGridRange(), declaration);
    return this;
  }
  addEditWarningFromRow(
    startRowIndex: number,
    declaration: EditWarningDeclaration = {},
  ): this {
    this.sheet.addEditWarningAt(
      this.gridRangeFromRow(startRowIndex),
      declaration,
    );
    return this;
  }
  addEditWarningWholeColumn(declaration: EditWarningDeclaration = {}): this {
    this.sheet.addEditWarningAt(this._wholeColumnGridRange(), declaration);
    return this;
  }
  addEditLock(declaration: EditLockDeclaration = {}): this {
    this.sheet.addEditLockAt(this.dataGridRange(), declaration);
    return this;
  }
  addEditLockWholeColumn(declaration: EditLockDeclaration = {}): this {
    this.sheet.addEditLockAt(this._wholeColumnGridRange(), declaration);
    return this;
  }
  removeEditProtections(): this {
    this.sheet.removeEditProtectionsAt(this.dataGridRange());
    return this;
  }
  removeEditProtectionsWholeColumn(): this {
    this.sheet.removeEditProtectionsAt(this._wholeColumnGridRange());
    return this;
  }
  removeEditProtection(protection: EditProtection): this {
    this.sheet.removeEditProtection(protection);
    return this;
  }
  gatherFetchActive(): this {
    this.cellIndexesActive.forEach((rowIndex) => {
      this.cell(rowIndex).gatherFetchRange();
    });
    return this;
  }
  gatherFetchFull(): this {
    const origin = this.tableOrigin();
    this.sheet.gatherFetchRange({
      startRowIndex: origin.sheetRowIndex(0),
      startColumnIndex: origin.sheetColIndex(this.colIndex),
      endColumnIndex: origin.sheetColIndex(this.colIndex + 1),
    });
    this.sheetState.fetchQueue.toFinalize.columns.add(this.colIndex);
    return this;
  }
  // A full-column fetch can hit rows that are entirely blank across every
  // column, which Sheets omits from the response — ensureStateExists
  // backfills those before ensureActive tries to touch a cell in them.
  ensureFullActiveDataCells(): void {
    this.sheet.rowIndexesFull.forEach((rowIndex) => {
      this.sheet.row(rowIndex).ensureStateExists();
      this.cell(rowIndex).ensureActive();
    });
  }
  private _wholeColumnGridRange(): GridRangeProps {
    return this._gridRangeFromSheetRow(SheetIndex.row(0));
  }
  private _gridRangeFromSheetRow(startRowIndex: SheetRowIndex): GridRangeProps {
    const origin = this.tableOrigin();
    return {
      sheetId: this.sheetGid,
      startRowIndex,
      startColumnIndex: origin.sheetColIndex(this.colIndex),
      endColumnIndex: origin.sheetColIndex(this.colIndex + 1),
    };
  }
}
