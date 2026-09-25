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
import { Arr } from "../utils/Arr";
import { CellRaw, validateFormulaString } from "./CellRaw";
import { ColumnBaseRaw } from "./ClassBases/ColumnBaseRaw";
import type { FindReplaceTerms, RowCellChange } from "./ClassTypes/StateRaw";
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
    return this.cell(this.schema.topDataRowIdx);
  }
  get dataGridRange(): BoundedGridRange {
    return {
      sheetId: this.sheetGid,
      startRowIndex: this.schema.topDataRowIdx,
      endRowIndex: this.sheet.activeTable.endRowIndex,
      startColumnIndex: this.colIndex,
      endColumnIndex: this.colIndex + 1,
    };
  }
  gridRangeFromRow(startRowIndex: number): GridRangeProps {
    return {
      sheetId: this.sheetGid,
      startRowIndex,
      startColumnIndex: this.colIndex,
      endColumnIndex: this.colIndex + 1,
    };
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
  updateAllCells(change: Omit<RowCellChange<VN>, "formula">): this {
    this.sheet.activeTable.assertRowIndexesNotStale();
    this.sheet.activeTable.validateColIndexNotStale(this.colIndex);
    this.sheet.validateNotPrunedToSelection();
    const { endRowIndex } = this.sheet.activeTable;
    const { value } = change;
    this.sheet.rowIndexesFull.forEach((rowIndex) => {
      const row = this.sheet.row(rowIndex);
      row.validateIsWritable();
      if (value !== undefined && row.rowIsActive()) {
        this.cell(rowIndex).setValueState(value);
      }
    });
    this.sheet.addSheetChangeToSave({
      action: "fill",
      colIndex: this.colIndex,
      startRowIndex: this.schema.topDataRowIdx,
      endRowIndex,
      ...change,
    });
    return this;
  }
  updateAllFormulas(formula: string): this {
    this.sheet.activeTable.assertRowIndexesNotStale();
    validateFormulaString(formula);
    this.sheet.activeTable.validateColIndexNotStale(this.colIndex);
    this.sheet.validateNotPrunedToSelection();
    const { endRowIndex } = this.sheet.activeTable;
    this.sheet.rowIndexesFull.forEach((rowIndex) => {
      this.sheet.row(rowIndex).validateIsWritable();
    });
    this.sheet.addSheetChangeToSave({
      action: "fill",
      colIndex: this.colIndex,
      startRowIndex: this.schema.topDataRowIdx,
      endRowIndex,
      formula,
    });
    return this;
  }
  updateActiveCells(change: Omit<RowCellChange<VN>, "formula">): this {
    this.sheet.activeTable.assertRowIndexesNotStale();
    this.sheet.activeTable.validateColIndexNotStale(this.colIndex);
    const rowIndexes = this.cellIndexesActive;
    const { value } = change;
    if (value !== undefined) {
      rowIndexes.forEach((rowIndex) => {
        this.cell(rowIndex).setValueState(value);
      });
    }
    Arr.contiguousRanges(rowIndexes).forEach(({ startIndex, endIndex }) => {
      this.sheet.addSheetChangeToSave({
        action: "fill",
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
    this.sheet.activeTable.validateColIndexNotStale(this.colIndex);
    Arr.contiguousRanges(this.cellIndexesActive).forEach(
      ({ startIndex, endIndex }) => {
        this.sheet.addSheetChangeToSave({
          action: "fill",
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
    this.sheet.activeTable.validateColIndexNotStale(this.colIndex);
    this.sheet.validateNotPrunedToSelection();
    this.ss.findReplace({ ...terms, scope: { range: this.dataGridRange } });
    return this;
  }
  addConditionalFormatRule(declaration: ConditionalFormatDeclaration): this {
    this.sheet.addConditionalFormatRuleAt(this.dataGridRange, declaration);
    return this;
  }
  removeConditionalFormatRules(): this {
    this.sheet.removeConditionalFormatRulesAt(this.dataGridRange);
    return this;
  }
  removeConditionalFormatRule(rule: ConditionalFormatRule): this {
    this.sheet.removeConditionalFormatRule(rule);
    return this;
  }
  addEditWarning(declaration: EditWarningDeclaration = {}): this {
    this.sheet.addEditWarningAt(this.dataGridRange, declaration);
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
    this.sheet.addEditWarningAt(this.gridRangeFromRow(0), declaration);
    return this;
  }
  addEditLock(declaration: EditLockDeclaration = {}): this {
    this.sheet.addEditLockAt(this.dataGridRange, declaration);
    return this;
  }
  addEditLockWholeColumn(declaration: EditLockDeclaration = {}): this {
    this.sheet.addEditLockAt(this.gridRangeFromRow(0), declaration);
    return this;
  }
  removeEditProtections(): this {
    this.sheet.removeEditProtectionsAt(this.dataGridRange);
    return this;
  }
  removeEditProtectionsWholeColumn(): this {
    this.sheet.removeEditProtectionsAt(this.gridRangeFromRow(0));
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
    this.sheet.gatherFetchRange({
      startRowIndex: this.schema.topDataRowIdx,
      startColumnIndex: this.colIndex,
      endColumnIndex: this.colIndex + 1,
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
}
