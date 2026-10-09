import { Val } from "@byronbroughten/utils/val";

import { SpreadsheetSchema } from "../../01_SpreadsheetSchema/configReaders/SpreadsheetSchema";
import type { TableName } from "../../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { headerZone } from "../../01_SpreadsheetSchema/headerZone";
import { SpreadsheetBaseRaw } from "../ClassBases/SpreadsheetBaseRaw";
import type { SheetStateRaw, TableStateRaw } from "../ClassTypes/StateRaw";
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
  // One per recorded Table the run gathered a fetch for; a broken one it doesn't use stays silent.
  tablePlacements(sheetGid: number): TablePlacement[] {
    if (this.spreadsheetStateRaw.isRegeneratingConfigs) return [];
    const { gatherHeaderZone } = this._sheetState(sheetGid).fetchQueue;
    return this.schema
      .tablesOnGid(sheetGid)
      .filter(({ tableName }) => this._hasGatheredFetch(tableName))
      .map(({ tableName }) => this.tablePlacement(tableName, gatherHeaderZone));
  }
  // What a run gathered waits on the live Table, or, before it is known, on its ID or its sheet.
  private _hasGatheredFetch(tableName: TableName): boolean {
    const table = this.schema.sheetByName(tableName);
    const sheetState = this._sheetState(table.sheetGid);
    const tableId = this._liveTableIdOf(tableName);
    const liveTableState =
      tableId === undefined
        ? undefined
        : this.spreadsheetStateRaw.tables.get(tableId);
    // A fetch through a shared sheet reaches none of its Tables, so it uses none of them.
    const sheetQueueState = table.sharesSheet
      ? undefined
      : sheetState.tableBeforeProperties;
    return [
      liveTableState,
      sheetState.tablesBeforePropertiesById.get(table.tableId),
      sheetQueueState,
    ].some(
      (tableState) => tableState !== undefined && hasGatheredFetch(tableState),
    );
  }
  private _sheetState(sheetGid: number): SheetStateRaw {
    return Val.assert(
      this.spreadsheetStateRaw.sheets.get(sheetGid),
      `sheetState for sheetGid ${sheetGid}`,
    );
  }
  // Missing and the column ID row wait for the zone, the one fetch sure to bring the Table.
  tablePlacement(tableName: TableName, isZoneFetched: boolean): TablePlacement {
    const table = this.schema.sheetByName(tableName);
    const { sheetGid } = table;
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
  // A Table is known only by its recorded ID, so one inserted again reads as missing.
  private _liveTableIdOf(tableName: TableName): string | undefined {
    const table = this.schema.sheetByName(tableName);
    const sheetTable = this.ss.tableOnSheet(table.sheetGid);
    if (sheetTable.tableIds().includes(table.tableId)) return table.tableId;
    return undefined;
  }
  validateTablePlacements(placements: TablePlacement[]): void {
    const fix = this._placementsFix(placements);
    if (fix === undefined) return;
    throw new Error(fix);
  }
  // A column insert can split a Table the zone missed from its head rows, so it uses every Table on its sheet.
  validateTablesForColumnInsert(sheetGid: number, columnInsert: string): void {
    if (this.spreadsheetStateRaw.isRegeneratingConfigs) return;
    const { hasFetchedHeaderZone } = this._sheetState(sheetGid).working;
    const fix = this._placementsFix(
      this.schema
        .tablesOnGid(sheetGid)
        .map(({ tableName }) =>
          this.tablePlacement(tableName, hasFetchedHeaderZone),
        ),
    );
    if (fix === undefined) return;
    throw new Error(
      `${columnInsert} needs every managed Table on that sheet in place. ${fix}`,
    );
  }
  private _placementsFix(placements: TablePlacement[]): string | undefined {
    const misplacements = placements.flatMap((placement) =>
      placement.kind === "misplaced" ? [placement.misplacement] : [],
    );
    const headerOnlyTableIds = placements.flatMap((placement) =>
      placement.kind === "header-only" ? [placement.tableId] : [],
    );
    if (misplacements.length === 0 && headerOnlyTableIds.length === 0) {
      return undefined;
    }
    const sentences: string[] = [];
    if (misplacements.length > 0) {
      sentences.push(this._misplacementsSentence(misplacements));
    }
    headerOnlyTableIds.forEach((tableId) => {
      sentences.push(this.ss.table(tableId).headerOnlyFix);
    });
    return sentences.join(" ");
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
    return `${misplacements.length} managed Table(s) are missing or misplaced: ${reasons}`;
  }
  // Generation refuses a Table outside the zone, so only the others are offered regeneration.
  private _misplacementReason(misplacement: Misplacement): string {
    if (misplacement.kind === "missing") {
      const name = this.schema
        .sheetByName(misplacement.tableName)
        .trait("tableName");
      return `has no Table "${name}" with ${headerRowPlace()} — move it back, or ${regenerateFix} if it is gone`;
    } else if (misplacement.kind === "outside-zone") {
      const { name } = this.ss.table(misplacement.tableId);
      return `has Table "${name}", which must have ${headerRowPlace()} — move it back`;
    } else if (misplacement.kind === "band-shifted") {
      const { idPrefix } = this.schema.sheetByName(misplacement.tableName);
      const colIdRowLabel = this.ss
        .table(misplacement.tableId)
        .rowLabel(this.schema.colIdRowIndex);
      return `needs its own "${idPrefix}" column IDs, and only those, in ${colIdRowLabel} — move the Table back, or ${regenerateFix}`;
    } else {
      throw new Error(`Unknown misplacement ${JSON.stringify(misplacement)}.`);
    }
  }
  private _sheetLabel({ sheetGid }: RecordedTableIdentity): string {
    return this.ss.sheet(sheetGid).label;
  }
}

function hasGatheredFetch({ fetchQueue }: TableStateRaw): boolean {
  const { rows, columns, cells } = fetchQueue.toFinalize;
  return rows.size > 0 || columns.size > 0 || cells.size > 0;
}

const regenerateFix =
  "regenerate the configs with sheets-framework gen-configs";

function headerZoneFix(tableLabel: string): string {
  return `${tableLabel} must have ${headerRowPlace()}.`;
}

function headerRowPlace(): string {
  return `its header row on ${headerZone.headerRowsLabel}`;
}
