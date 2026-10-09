import { SpreadsheetSchema } from "../../01_SpreadsheetSchema/configReaders/SpreadsheetSchema";
import type { TableSchema } from "../../01_SpreadsheetSchema/configReaders/TableSchema";
import { headerZone } from "../../01_SpreadsheetSchema/headerZone";
import { SpreadsheetRaw } from "../../02_SpreadsheetRaw/SpreadsheetRaw";
import { headerRowPlace } from "../../02_SpreadsheetRaw/SpreadsheetRaw/SpreadsheetTableValidatorRaw";
import { SpreadsheetBaseIdentified } from "../ClassBases/SpreadsheetBaseIdentified";

type Misplacement = { table: TableSchema } & (
  { kind: "missing" } | { kind: "outside-zone" } | { kind: "band-shifted" }
);

export class SpreadsheetTableValidatorIdentified extends SpreadsheetBaseIdentified {
  get raw(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
  get schema(): SpreadsheetSchema {
    return new SpreadsheetSchema();
  }
  // One per recorded Table the last fetch gathered for; a broken one the run doesn't use stays silent.
  validateUsedTables(): void {
    if (this.spreadsheetStateIdentified.isRegeneratingConfigs) return;
    const { usedTableIds } = this.raw;
    this.validateTables(
      this.raw.activeSheetGids
        .flatMap((sheetGid) => this.schema.tablesOnGid(sheetGid))
        .filter(({ tableId }) => usedTableIds.includes(tableId)),
    );
  }
  validateTables(tables: TableSchema[]): void {
    const fix = this._placementFix(tables);
    if (fix === undefined) return;
    throw new Error(fix);
  }
  // A column insert can split a Table the zone missed from its head rows, so it uses every Table on its sheet.
  validateTablesForColumnInserts(): void {
    if (this.spreadsheetStateIdentified.isRegeneratingConfigs) return;
    for (const tableRaw of this.raw.tablesWithQueuedColumnInsert) {
      const fix = this._placementFix(
        this.schema.tablesOnGid(tableRaw.sheetGid),
      );
      if (fix === undefined) continue;
      throw new Error(
        `${tableRaw.columnInsertLabel} needs every managed Table on that sheet in place. ${fix}`,
      );
    }
  }
  private _placementFix(tables: TableSchema[]): string | undefined {
    const headerOnlyFix = this._headerOnlyFix(tables);
    if (headerOnlyFix !== undefined) return headerOnlyFix;
    const misplacements = tables.flatMap((table) => this._misplacements(table));
    if (misplacements.length === 0) return undefined;
    return this._misplacementsSentence(misplacements);
  }
  // First and alone, since the band test reads the column ID row through the body origin.
  private _headerOnlyFix(tables: TableSchema[]): string | undefined {
    const fixes = tables
      .filter((table) => this._isLive(table))
      .map(({ tableId }) => this.raw.table(tableId))
      .filter((tableRaw) => tableRaw.isHeaderOnly)
      .map((tableRaw) => tableRaw.headerOnlyFix);
    if (fixes.length === 0) return undefined;
    return fixes.join(" ");
  }
  // Missing and the column ID row wait for the zone, the one fetch sure to bring the Table.
  private _misplacements(table: TableSchema): Misplacement[] {
    const { hasFetchedHeaderZone } = this.raw.sheet(table.sheetGid);
    if (!this._isLive(table)) {
      return hasFetchedHeaderZone ? [{ kind: "missing", table }] : [];
    }
    const tableRaw = this.raw.table(table.tableId);
    if (!headerZone.holdsHeaderRow(tableRaw.startRowIndex)) {
      return [{ kind: "outside-zone", table }];
    }
    if (
      hasFetchedHeaderZone &&
      !tableRaw.columnResolver.holdsOnlyColumnIdsOf(table.idPrefix)
    ) {
      return [{ kind: "band-shifted", table }];
    }
    return [];
  }
  // A Table is known only by its recorded ID, so one inserted again reads as missing.
  private _isLive({ sheetGid, tableId }: TableSchema): boolean {
    return this.raw.sheet(sheetGid).tableIds.includes(tableId);
  }
  private _misplacementsSentence(misplacements: Misplacement[]): string {
    const reasons = misplacements
      .map(
        (misplacement) =>
          `${this.raw.sheet(misplacement.table.sheetGid).label} ${this._misplacementReason(misplacement)}`,
      )
      .join("; ");
    return `${misplacements.length} managed Table(s) are missing or misplaced: ${reasons}`;
  }
  // Generation refuses a Table outside the zone, so only the others are offered regeneration.
  private _misplacementReason({ kind, table }: Misplacement): string {
    if (kind === "missing") {
      return `has no Table "${table.trait("tableName")}" with ${headerRowPlace()} — move it back, or ${regenerateFix} if it is gone`;
    } else if (kind === "outside-zone") {
      const { name } = this.raw.table(table.tableId);
      return `has Table "${name}", which must have ${headerRowPlace()} — move it back`;
    } else if (kind === "band-shifted") {
      const colIdRowLabel = this.raw
        .table(table.tableId)
        .rowLabel(this.schema.colIdRowIndex);
      return `needs its own "${table.idPrefix}" column IDs, and only those, in ${colIdRowLabel} — move the Table back, or ${regenerateFix}`;
    } else {
      const exhaustive: never = kind;
      throw new Error(`Unknown misplacement ${JSON.stringify(exhaustive)}.`);
    }
  }
}

const regenerateFix =
  "regenerate the configs with sheets-framework gen-configs";
