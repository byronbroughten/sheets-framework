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
  TableColumnType,
} from "../00_Source/RawSource/RawSource";
import {
  SheetIndex,
  type SheetRowIndex,
} from "../00_Source/RawSource/SheetIndex";
import {
  type HeadRole,
  headRows,
  type HeadRowValueName,
} from "../01_SpreadsheetSchema/headRows";
import { Arr } from "../utils/Arr";
import { CellRaw, validateFormulaString } from "./CellRaw";
import { ColumnBaseRaw } from "./ClassBases/ColumnBaseRaw";
import type {
  CellFill,
  FindReplaceTerms,
  TableEndColumnHeadCells,
} from "./ClassTypes/StateRaw";
import { ColumnProfileRaw } from "./ColumnProfileRaw";
import { SpreadsheetRaw } from "./SpreadsheetRaw";
import { TableRaw } from "./TableRaw";

export class ColumnRaw<
  VN extends CellValueName = CellValueName,
> extends ColumnBaseRaw {
  get ss(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
  get table(): TableRaw {
    return new TableRaw(this.tableRawProps);
  }
  get profile(): ColumnProfileRaw {
    return new ColumnProfileRaw(this.columnRawProps);
  }
  // The live names, so formula text stays right after a rename without regenerating.
  get reference(): string {
    const { header } = this.profile;
    if (header === "") {
      throw new Error(
        `Column ${this.colIndex} of ${this.table.tableLabel} has a blank header, so it has no Table reference.`,
      );
    }
    return `${this.table.name}[${header}]`;
  }
  get single(): string {
    return `SINGLE(${this.reference})`;
  }
  get valueArrOrEmpty(): (CellValue<VN> | "")[] {
    return this.table.workingRowIndexes.map((rowIndex) =>
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
    const { origin, dataRowCount } = this.table;
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
  get workingCellIndexes(): number[] {
    return this.table.workingRowIndexes;
  }
  get cellIndexesFull(): number[] {
    return this.table.rowIndexesFull;
  }
  cell(rowIndex: number): CellRaw<VN> {
    return new CellRaw<VN>({
      ...this.columnRawProps,
      rowIndex,
    });
  }
  headCell<HR extends HeadRole>(headRole: HR): CellRaw<HeadRowValueName<HR>> {
    return new CellRaw<HeadRowValueName<HR>>({
      ...this.columnRawProps,
      rowIndex: headRows.index(headRole),
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
    this.table.assertRowIndexesNotStale();
    this.table.validateNotPrunedToSelection();
    const { dataRowCount } = this.table;
    const { value } = change;
    this.table.rowIndexesFull.forEach((rowIndex) => {
      const row = this.table.row(rowIndex);
      row.validateIsWritable();
      if (value !== undefined && row.rowInWorking()) {
        this.cell(rowIndex).setValueState(value);
      }
    });
    this.table.queueTableWrite({
      action: "fillColumn",
      colIndex: this.colIndex,
      startRowIndex: 0,
      endRowIndex: dataRowCount,
      ...change,
    });
    return this;
  }
  updateAllFormulas(formula: string): this {
    this.table.assertRowIndexesNotStale();
    validateFormulaString(formula);
    this.table.validateNotPrunedToSelection();
    const { dataRowCount } = this.table;
    this.table.rowIndexesFull.forEach((rowIndex) => {
      this.table.row(rowIndex).validateIsWritable();
    });
    this.table.queueTableWrite({
      action: "fillColumn",
      colIndex: this.colIndex,
      startRowIndex: 0,
      endRowIndex: dataRowCount,
      formula,
    });
    return this;
  }
  updateWorkingCells(change: Omit<CellFill<VN>, "formula">): this {
    this.table.assertRowIndexesNotStale();
    const rowIndexes = this.workingCellIndexes;
    const { value } = change;
    if (value !== undefined) {
      rowIndexes.forEach((rowIndex) => {
        this.cell(rowIndex).setValueState(value);
      });
    }
    Arr.contiguousRanges(rowIndexes).forEach(({ startIndex, endIndex }) => {
      this.table.queueTableWrite({
        action: "fillColumn",
        colIndex: this.colIndex,
        startRowIndex: startIndex,
        endRowIndex: endIndex,
        ...change,
      });
    });
    return this;
  }
  updateWorkingFormulas(formula: string): this {
    this.table.assertRowIndexesNotStale();
    validateFormulaString(formula);
    Arr.contiguousRanges(this.workingCellIndexes).forEach(
      ({ startIndex, endIndex }) => {
        this.table.queueTableWrite({
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
    this.table.assertRowIndexesNotStale();
    this.table.validateNotPrunedToSelection();
    this.table.queueTableWrite({
      action: "findReplace",
      terms,
      startColIndex: this.colIndex,
      endColIndex: this.colIndex + 1,
    });
    return this;
  }
  // Called only by the Table's appendColumn, on the column it just queued.
  _initHeadCells({
    columnId,
    header,
    groupHeading1,
  }: TableEndColumnHeadCells): this {
    this.headCell("columnId").updateValue(columnId);
    this.headCell("header").updateValue(header);
    if (groupHeading1 !== undefined) {
      this.headCell("groupHeading1").updateValue(groupHeading1);
    }
    return this;
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
  addConditionalFormatRule(declaration: ConditionalFormatDeclaration): this {
    this.table.sheet.addConditionalFormatRuleAt(
      this.dataGridRange(),
      declaration,
    );
    return this;
  }
  removeConditionalFormatRules(): this {
    this.table.sheet.removeConditionalFormatRulesAt(this.dataGridRange());
    return this;
  }
  removeConditionalFormatRule(rule: ConditionalFormatRule): this {
    this.table.sheet.removeConditionalFormatRule(rule);
    return this;
  }
  addEditWarning(declaration: EditWarningDeclaration = {}): this {
    this.table.sheet.addEditWarningAt(this.dataGridRange(), declaration);
    return this;
  }
  addEditWarningFromRow(
    startRowIndex: number,
    declaration: EditWarningDeclaration = {},
  ): this {
    this.table.sheet.addEditWarningAt(
      this.gridRangeFromRow(startRowIndex),
      declaration,
    );
    return this;
  }
  addEditWarningWholeColumn(declaration: EditWarningDeclaration = {}): this {
    this.table.sheet.addEditWarningAt(
      this._wholeColumnGridRange(),
      declaration,
    );
    return this;
  }
  addEditLock(declaration: EditLockDeclaration = {}): this {
    this.table.sheet.addEditLockAt(this.dataGridRange(), declaration);
    return this;
  }
  addEditLockWholeColumn(declaration: EditLockDeclaration = {}): this {
    this.table.sheet.addEditLockAt(this._wholeColumnGridRange(), declaration);
    return this;
  }
  removeEditProtections(): this {
    this.table.sheet.removeEditProtectionsAt(this.dataGridRange());
    return this;
  }
  removeEditProtectionsWholeColumn(): this {
    this.table.sheet.removeEditProtectionsAt(this._wholeColumnGridRange());
    return this;
  }
  removeEditProtection(protection: EditProtection): this {
    this.table.sheet.removeEditProtection(protection);
    return this;
  }
  gatherFetchWorking(): this {
    this.workingCellIndexes.forEach((rowIndex) => {
      this.cell(rowIndex).gatherFetchRange();
    });
    return this;
  }
  gatherFetchFull(): this {
    this.table.gatherFetchRange(this.table.fullColumnFetchRange(this.colIndex));
    this.tableState.fetchQueue.toFinalize.columns.add(this.colIndex);
    return this;
  }
  // A full-column fetch can hit rows that are entirely blank across every
  // column, which Sheets omits from the response — ensureStateExists
  // backfills those before ensureInWorking tries to touch a cell in them.
  ensureFullWorkingDataCells(): void {
    this.table.rowIndexesFull.forEach((rowIndex) => {
      this.table.row(rowIndex).ensureStateExists();
      this.cell(rowIndex).ensureInWorking();
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
