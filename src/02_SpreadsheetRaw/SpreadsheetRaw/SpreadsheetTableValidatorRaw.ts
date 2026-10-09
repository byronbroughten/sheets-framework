import { headerZone } from "../../01_SpreadsheetSchema/headerZone";
import { SpreadsheetBaseRaw } from "../ClassBases/SpreadsheetBaseRaw";
import { SpreadsheetRaw } from "../SpreadsheetRaw";

export class SpreadsheetTableValidatorRaw extends SpreadsheetBaseRaw {
  get ss(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
  // Generation's counterpart to the placement check, so configs never record a Table a run would stop on.
  validateHeadersInZone(managedTableIds: string[]): void {
    const sentences = managedTableIds
      .map((tableId) => this.ss.table(tableId))
      .filter((table) => !headerZone.holdsHeaderRow(table.startRowIndex))
      .map((table) => headerZoneFix(table.tableLabel));
    if (sentences.length === 0) return;
    throw new Error(sentences.join(" "));
  }
  // By geometry alone: the column ID row may give way to metadata on the header row.
  validateHeadRowsClear(managedTableIds: string[]): void {
    const sentences = managedTableIds.flatMap((tableId) => {
      const table = this.ss.table(tableId);
      return table
        .tableIds()
        .filter((otherId) => otherId !== tableId)
        .map((otherId) => this.ss.table(otherId))
        .filter((other) => table.headRowsSitOn(other))
        .map((other) => table.headRowsOverlapFix(other));
    });
    if (sentences.length === 0) return;
    throw new Error(sentences.join(" "));
  }
}

function headerZoneFix(tableLabel: string): string {
  return `${tableLabel} must have ${headerRowPlace()}.`;
}

export function headerRowPlace(): string {
  return `its header row on ${headerZone.headerRowsLabel}`;
}
