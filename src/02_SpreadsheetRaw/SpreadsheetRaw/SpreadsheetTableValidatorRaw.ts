import type {
  SheetColIndex,
  SheetRowIndex,
} from "../../00_Source/RawSource/SheetIndex";
import { SpreadsheetSchema } from "../../01_SpreadsheetSchema/SpreadsheetSchema";
import { Val } from "../../utils/Val";
import { SpreadsheetBaseRaw } from "../ClassBases/SpreadsheetBaseRaw";
import { SpreadsheetRaw } from "../SpreadsheetRaw";

interface SheetIdentity {
  sheetGid: number;
}
export type Misplacement = SheetIdentity &
  (
    | { kind: "missing" }
    | {
        kind: "moved";
        startRowIndex: SheetRowIndex;
        startColumnIndex: SheetColIndex;
      }
    | { kind: "band-shifted" }
  );
export type TablePlacement =
  | { kind: "extra" }
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
  // A sheet outside the config never promised to follow the layout.
  tablePlacement(sheetGid: number): TablePlacement {
    const sheetState = this.spreadsheetStateRaw.sheets.get(sheetGid);
    if (sheetState === undefined) {
      return { kind: "none" };
    }
    const [tableId, ...otherTableIds] = this.ss.sheet(sheetGid).tableIds();
    if (otherTableIds.length > 0) {
      return { kind: "extra" };
    }
    if (!this.schema.isInSheetGids(sheetGid)) {
      return { kind: "none" };
    }
    const isStripFetched = sheetState.fetchQueue.gatherPlacementStrip;
    if (tableId === undefined && !isStripFetched) {
      return { kind: "none" };
    }
    if (tableId === undefined) {
      return { kind: "misplaced", misplacement: { kind: "missing", sheetGid } };
    }
    // Read off the state, since the Table refuses a header-only body before placement is judged.
    const { startRowIndex, startColumnIndex } = Val.assert(
      this.spreadsheetStateRaw.tables.get(tableId)?.properties,
      `properties of Table ${tableId}`,
    );
    if (!this.schema.isTableStart(startRowIndex, startColumnIndex)) {
      return {
        kind: "misplaced",
        misplacement: { kind: "moved", sheetGid, startRowIndex, startColumnIndex },
      };
    }
    // Before the band test, which reads the column ID row through the Table's body origin.
    if (this.ss.table(tableId).isHeaderOnly) {
      return { kind: "header-only", tableId };
    }
    if (isStripFetched && !this._holdsOwnColumnIds(sheetGid)) {
      return {
        kind: "misplaced",
        misplacement: { kind: "band-shifted", sheetGid },
      };
    }
    return { kind: "well-placed" };
  }
  validateTablePlacement(
    misplacements: Misplacement[] = [],
    headerOnlyTableIds: string[] = [],
  ): void {
    const extraTables = this._sheetsWithExtraTables();
    if (
      misplacements.length === 0 &&
      extraTables.length === 0 &&
      headerOnlyTableIds.length === 0
    ) {
      return;
    }
    const sentences: string[] = [];
    if (misplacements.length > 0) {
      sentences.push(this._misplacementsSentence(misplacements));
    }
    if (extraTables.length > 0) {
      sentences.push(this._extraTablesSentence(extraTables));
    }
    headerOnlyTableIds.forEach((tableId) => {
      sentences.push(this.ss.table(tableId).headerOnlyFix);
    });
    throw new Error(sentences.join(" "));
  }
  private _holdsOwnColumnIds(sheetGid: number): boolean {
    const { idPrefix } = this.schema.sheetByGid(sheetGid);
    return this.ss.sheetMeta(sheetGid).holdsOnlyColumnIdsOf(idPrefix);
  }
  private _sheetsWithExtraTables(): SheetIdentity[] {
    const extraTables: SheetIdentity[] = [];
    this.spreadsheetStateRaw.sheets.forEach((_, sheetGid) => {
      if (
        this.ss.sheet(sheetGid).tableIds().length <= 1 ||
        !this.schema.isInSheetGids(sheetGid)
      ) {
        return;
      }
      extraTables.push({ sheetGid });
    });
    return extraTables;
  }
  private _misplacementsSentence(misplacements: Misplacement[]): string {
    const reasons = misplacements
      .map(
        (misplacement) =>
          `${this._sheetLabel(misplacement)} ${this._misplacementReason(misplacement)}`,
      )
      .join("; ");
    return `${misplacements.length} managed Table(s) are not where the configs record them — regenerate the configs with sheets-framework gen-configs: ${reasons}`;
  }
  private _misplacementReason(misplacement: Misplacement): string {
    if (misplacement.kind === "missing") {
      return `has no Table starting at ${this.schema.tableStartLabel}`;
    } else if (misplacement.kind === "moved") {
      return `has a Table that starts at ${this.schema.positionLabel(
        misplacement.startRowIndex,
        misplacement.startColumnIndex,
      )}, not ${this.schema.tableStartLabel}`;
    } else if (misplacement.kind === "band-shifted") {
      const { idPrefix } = this.schema.sheetByGid(misplacement.sheetGid);
      const colIdRowLabel = this.ss
        .sheetMeta(misplacement.sheetGid)
        .rowLabel(this.schema.colIdRowIndex);
      return `needs its own "${idPrefix}" column IDs, and only those, in ${colIdRowLabel}`;
    } else {
      throw new Error(`Unknown misplacement ${JSON.stringify(misplacement)}.`);
    }
  }
  private _extraTablesSentence(extraTables: SheetIdentity[]): string {
    const names = extraTables
      .map((extraTable) => this._sheetLabel(extraTable))
      .join(", ");
    return `${extraTables.length} sheet(s) have more than one Table — delete the extras so each sheet has exactly one: ${names}`;
  }
  private _sheetLabel({ sheetGid }: SheetIdentity): string {
    return this.ss.sheet(sheetGid).sheetLabel;
  }
}
