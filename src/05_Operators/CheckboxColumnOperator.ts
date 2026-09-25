import type { ColumnNameFiltered } from "../01_SpreadsheetSchema/columnConfigsTypes";
import type { SheetNameSimple } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import { ColumnIdentified } from "../03_SpreadsheetIdentified/ColumnIdentified";
import { ColumnBaseNamed } from "../04_SpreadsheetNamed/ClassBases/ColumnBaseNamed";
import { ColumnNamed } from "../04_SpreadsheetNamed/ColumnNamed";
import type { SheetNamed } from "../04_SpreadsheetNamed/SheetNamed";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";

// `checkbox`, not `boolean`: only a declared checkbox column is never blank.
export type CheckboxColumnName<SN extends SheetNameSimple> = ColumnNameFiltered<
  SN,
  "checkbox",
  false
>;

export class CheckboxColumnOperator<
  SN extends SheetNameSimple,
  CN extends CheckboxColumnName<SN>,
> extends ColumnBaseNamed<SN, CN> {
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  get sheet(): SheetNamed<SN> {
    return this.ss.sheet(this.sheetName);
  }
  get column(): ColumnNamed<SN, CN> {
    return new ColumnNamed(this.columnNamedProps);
  }
  // Named can't re-derive `checkbox` while SN is generic, so the write is pinned here.
  get identified(): ColumnIdentified<"checkbox"> {
    return new ColumnIdentified<"checkbox">({
      ...this.sheet.identified.sheetIdentifiedProps,
      columnId: this.column.columnId,
    });
  }
  get rowIndexesChecked(): number[] {
    return this.column.rowIndexesActive.filter(
      (rowIndex) => this.column.value(rowIndex) === true,
    );
  }
  // The active-cells fill, so an uncheck is safe on a sheet pruned to a selection.
  uncheckActiveCells(): this {
    this.identified.updateActiveCells({ value: false });
    return this;
  }
}
