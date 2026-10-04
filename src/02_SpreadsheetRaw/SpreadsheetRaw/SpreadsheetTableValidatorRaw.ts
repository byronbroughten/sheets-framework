import type {
  SheetColIndex,
  SheetRowIndex,
} from "../../00_Source/RawSource/SheetIndex";
import { SpreadsheetSchema } from "../../01_SpreadsheetSchema/SpreadsheetSchema";
import { Val } from "../../utils/Val";
import { SpreadsheetBaseRaw } from "../ClassBases/SpreadsheetBaseRaw";
import { SpreadsheetRaw } from "../SpreadsheetRaw";

export interface SheetIdentity {
  sheetGid: number;
}
export interface MisplacedTable extends SheetIdentity {
  startRowIndex: SheetRowIndex;
  startColumnIndex: SheetColIndex;
}
interface TablePlacementObservations {
  misplacedTables: MisplacedTable[];
  absentTables: SheetIdentity[];
}
const noObservations: TablePlacementObservations = {
  misplacedTables: [],
  absentTables: [],
};
export type TablePlacement =
  | { kind: "extra" }
  | (MisplacedTable & { kind: "misplaced" })
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
    if (!this.spreadsheetStateRaw.sheets.has(sheetGid)) {
      return { kind: "none" };
    }
    const [tableId, ...otherTableIds] = this.ss.sheet(sheetGid).tableIds();
    if (otherTableIds.length > 0) {
      return { kind: "extra" };
    }
    if (tableId === undefined || !this.schema.isInSheetGids(sheetGid)) {
      return { kind: "none" };
    }
    // Read off the state, since the Table refuses a header-only body before placement is judged.
    const { startRowIndex, startColumnIndex } = Val.assert(
      this.spreadsheetStateRaw.tables.get(tableId)?.properties,
      `properties of Table ${tableId}`,
    );
    if (this.schema.isTableStart(startRowIndex, startColumnIndex)) {
      return { kind: "well-placed" };
    }
    return { kind: "misplaced", sheetGid, startRowIndex, startColumnIndex };
  }
  validateTablePlacement({
    misplacedTables,
    absentTables,
  }: TablePlacementObservations = noObservations): void {
    const reclassified = this._reclassifyAbsentTables({
      misplacedTables,
      absentTables,
    });
    const extraTables = this._sheetsWithExtraTables();
    if (
      reclassified.misplacedTables.length === 0 &&
      reclassified.absentTables.length === 0 &&
      extraTables.length === 0
    ) {
      return;
    }
    const sentences: string[] = [];
    if (reclassified.misplacedTables.length > 0) {
      sentences.push(
        this._misplacedTableSentence(reclassified.misplacedTables),
      );
    }
    if (reclassified.absentTables.length > 0) {
      sentences.push(this._absentTableSentence(reclassified.absentTables));
    }
    if (extraTables.length > 0) {
      sentences.push(this._extraTablesSentence(extraTables));
    }
    throw new Error(sentences.join(" "));
  }
  private _reclassifyAbsentTables({
    misplacedTables,
    absentTables,
  }: TablePlacementObservations): TablePlacementObservations {
    const stillMisplaced = [...misplacedTables];
    const stillAbsent: SheetIdentity[] = [];
    absentTables.forEach((absentTable) => {
      const placement = this.tablePlacement(absentTable.sheetGid);
      if (placement.kind === "extra") {
        return;
      }
      if (placement.kind === "misplaced") {
        stillMisplaced.push(placement);
        return;
      }
      stillAbsent.push(absentTable);
    });
    return { misplacedTables: stillMisplaced, absentTables: stillAbsent };
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
  private _misplacedTableSentence(misplacedTables: MisplacedTable[]): string {
    const positions = misplacedTables
      .map(
        (misplacedTable) =>
          `${this._sheetLabel(misplacedTable)} starts at ${this.schema.positionLabel(
            misplacedTable.startRowIndex,
            misplacedTable.startColumnIndex,
          )} but must start at ${this.schema.tableStartLabel}`,
      )
      .join("; ");
    return `${misplacedTables.length} sheet(s) have a Table that does not start where the layout requires — move each Table to where it must start, and do not rebuild it: ${positions}`;
  }
  private _absentTableSentence(absentTables: SheetIdentity[]): string {
    const names = absentTables
      .map((absentTable) => this._sheetLabel(absentTable))
      .join(", ");
    return `${absentTables.length} sheet(s) need a full row/column fetch but have no Table object — apply Insert > Table over their data range in Sheets: ${names}`;
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
