import type { ColumnNameFiltered } from "../01_SpreadsheetSchema/columnConfigsTypes";
import type { TableNameSimple } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import { ColumnIdentified } from "../03_SpreadsheetIdentified/ColumnIdentified";
import { ColumnBaseNamed } from "../04_SpreadsheetNamed/ClassBases/ColumnBaseNamed";
import { ColumnNamed } from "../04_SpreadsheetNamed/ColumnNamed";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import type { TableNamed } from "../04_SpreadsheetNamed/TableNamed";

// `checkbox`, not `boolean`: only a declared checkbox column is never blank.
export type CheckboxColumnName<TN extends TableNameSimple> = ColumnNameFiltered<
  TN,
  "checkbox",
  false
>;

export class CheckboxColumnOperator<
  TN extends TableNameSimple,
  CN extends CheckboxColumnName<TN>,
> extends ColumnBaseNamed<TN, CN> {
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  get sheet(): TableNamed<TN> {
    return this.ss.sheet(this.sheetName);
  }
  get column(): ColumnNamed<TN, CN> {
    return new ColumnNamed(this.columnNamedProps);
  }
  // Named can't re-derive `checkbox` while TN is generic, so the write is pinned here.
  get identified(): ColumnIdentified<"checkbox"> {
    return new ColumnIdentified<"checkbox">({
      ...this.sheet.identified.tableIdentifiedProps,
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
