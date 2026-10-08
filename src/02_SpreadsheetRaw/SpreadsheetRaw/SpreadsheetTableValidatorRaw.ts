import { headerZone } from "../../01_SpreadsheetSchema/headerZone";
import { SpreadsheetSchema } from "../../01_SpreadsheetSchema/SpreadsheetSchema";
import type { TableName } from "../../01_SpreadsheetSchema/tableConfigsTypes";
import { Val } from "../../utils/Val";
import { SpreadsheetBaseRaw } from "../ClassBases/SpreadsheetBaseRaw";
import { SpreadsheetRaw } from "../SpreadsheetRaw";

interface RecordedTableIdentity {
  sheetGid: number;
  tableName: TableName;
}
export type Misplacement = RecordedTableIdentity &
  (
    | { kind: "missing" }
    | { kind: "outside-zone"; tableId: string }
    | { kind: "band-shifted"; tableId: string }
  );
export type TablePlacement =
  | { kind: "header-only"; tableId: string }
  | { kind: "misplaced"; misplacement: Misplacement }
  | { kind: "none" }
  | { kind: "well-placed" };

export class SpreadsheetTableValidatorRaw extends SpreadsheetBaseRaw {
  get ss(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
  get schema(): SpreadsheetSchema {
    return new SpreadsheetSchema();
  }
  // One per Table the configs record on the sheet; a sheet outside the config never promised to follow the layout.
  tablePlacements(sheetGid: number): TablePlacement[] {
    return this.schema
      .tablesOnGid(sheetGid)
      .map(({ tableName }) => this._tablePlacement(tableName));
  }
  // Missing and the column ID row wait for the zone, the one fetch sure to bring the Table.
  private _tablePlacement(tableName: TableName): TablePlacement {
    const table = this.schema.sheetByName(tableName);
    const { sheetGid } = table;
    const isZoneFetched = Val.assert(
      this.spreadsheetStateRaw.sheets.get(sheetGid),
      `sheetState for sheetGid ${sheetGid}`,
    ).fetchQueue.gatherHeaderZone;
    const tableId = this._liveTableIdOf(tableName);
    if (tableId === undefined && !isZoneFetched) {
      return { kind: "none" };
    }
    if (tableId === undefined) {
      return {
        kind: "misplaced",
        misplacement: { kind: "missing", sheetGid, tableName },
      };
    }
    // Read off the state, since the Table refuses a header-only body before placement is judged.
    const { properties } = Val.assert(
      this.spreadsheetStateRaw.tables.get(tableId),
      `state of Table ${tableId}`,
    );
    const { startRowIndex } = Val.assert(
      properties,
      `properties of Table ${tableId}`,
    );
    if (!headerZone.holdsHeaderRow(startRowIndex)) {
      return {
        kind: "misplaced",
        misplacement: { kind: "outside-zone", sheetGid, tableName, tableId },
      };
    }
    // Before the band test, which reads the column ID row through the Table's body origin.
    if (this.ss.table(tableId).isHeaderOnly) {
      return { kind: "header-only", tableId };
    }
    if (isZoneFetched && !this._holdsOwnColumnIds(tableId, table.idPrefix)) {
      return {
        kind: "misplaced",
        misplacement: { kind: "band-shifted", sheetGid, tableName, tableId },
      };
    }
    return { kind: "well-placed" };
  }
  // On a sheet the configs record several Tables on, each is known only by its recorded ID.
  private _liveTableIdOf(tableName: TableName): string | undefined {
    const table = this.schema.sheetByName(tableName);
    const sheetTable = this.ss.tableOnSheet(table.sheetGid);
    if (sheetTable.tableIds().includes(table.tableId)) return table.tableId;
    if (table.sharesSheet) return undefined;
    return sheetTable.tableIdReachedByGid();
  }
  validateTablePlacement(
    misplacements: Misplacement[],
    headerOnlyTableIds: string[],
  ): void {
    if (misplacements.length === 0 && headerOnlyTableIds.length === 0) {
      return;
    }
    const sentences: string[] = [];
    if (misplacements.length > 0) {
      sentences.push(this._misplacementsSentence(misplacements));
    }
    headerOnlyTableIds.forEach((tableId) => {
      sentences.push(this.ss.table(tableId).headerOnlyFix);
    });
    throw new Error(sentences.join(" "));
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
  private _holdsOwnColumnIds(tableId: string, idPrefix: string): boolean {
    return this.ss.table(tableId).columnResolver.holdsOnlyColumnIdsOf(idPrefix);
  }
  private _misplacementsSentence(misplacements: Misplacement[]): string {
    const reasons = misplacements
      .map(
        (misplacement) =>
          `${this._sheetLabel(misplacement)} ${this._misplacementReason(misplacement)}`,
      )
      .join("; ");
    return `${misplacements.length} managed Table(s) are missing or outside the header zone — move each back, or regenerate the configs with sheets-framework gen-configs: ${reasons}`;
  }
  private _misplacementReason(misplacement: Misplacement): string {
    if (misplacement.kind === "missing") {
      const name = this.schema
        .sheetByName(misplacement.tableName)
        .trait("tableName");
      return `has no Table "${name}" with ${headerRowPlace()}`;
    } else if (misplacement.kind === "outside-zone") {
      const { name } = this.ss.table(misplacement.tableId);
      return `has Table "${name}", which must have ${headerRowPlace()}`;
    } else if (misplacement.kind === "band-shifted") {
      const { idPrefix } = this.schema.sheetByName(misplacement.tableName);
      const colIdRowLabel = this.ss
        .table(misplacement.tableId)
        .rowLabel(this.schema.colIdRowIndex);
      return `needs its own "${idPrefix}" column IDs, and only those, in ${colIdRowLabel}`;
    } else {
      throw new Error(`Unknown misplacement ${JSON.stringify(misplacement)}.`);
    }
  }
  private _sheetLabel({ sheetGid }: RecordedTableIdentity): string {
    return this.ss.sheet(sheetGid).label;
  }
}

function headerZoneFix(tableLabel: string): string {
  return `${tableLabel} must have ${headerRowPlace()}.`;
}

function headerRowPlace(): string {
  return `its header row on ${headerZone.headerRowsLabel}`;
}
