import {
  type ColumnName,
  getColumnTraitByName,
} from "../../01_SpreadsheetSchema/columnConfigsTypes";
import { getSheetTraitByName } from "../../01_SpreadsheetSchema/sheetConfigsTypes";
import { SheetBaseNamed } from "../../04_SpreadsheetNamed/ClassBases/SheetBaseNamed";
import type { SheetNamed } from "../../04_SpreadsheetNamed/SheetNamed";
import { SpreadsheetNamed } from "../../04_SpreadsheetNamed/SpreadsheetNamed";
import { liveColIndex } from "./floorColumnLocation";
import {
  columnNameByHeader,
  type FloorColumnRestore,
  floorColumnsToRestore,
  type FloorSheetName,
  spreadsheetConfigFeedbackColumnNames,
} from "./floorSeedLookups";

export class FloorTabColumnCreator<
  SN extends FloorSheetName,
> extends SheetBaseNamed<SN> {
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  get sheet(): SheetNamed<SN> {
    return this.ss.sheet(this.sheetName);
  }
  hasFloorTable(): boolean {
    const sheetGid = getSheetTraitByName(this.sheetName, "sheetGid");
    if (!this.ss.raw.gidIsActive(sheetGid)) return false;
    return this.sheet.raw.tables.length === 1;
  }
  assertMissingAreRecreatable(): void {
    this._assertTableMenuSpaceIsFirst();
    const recreatable = recreatableColumns()[this.sheetName];
    this._missingColumns().forEach((floorColumn) => {
      const columnName = columnNameByHeader(this.sheetName, floorColumn.header);
      if (recreatable.includes(columnName)) return;
      throw new Error(
        `${floorColumnLabel(floorColumn.header)} is missing from ${this.sheet.raw.title}, and recreating it empty would lose what it held. Undo the delete, or insert a column headed "${floorColumn.header}" in its Table and fill it.`,
      );
    });
  }
  // Putting it back first would be an insert inside the Table, which stales every index right of it.
  private _assertTableMenuSpaceIsFirst(): void {
    if (this.sheetName !== "spreadsheetConfig") return;
    const meta = this.sheet.raw.meta;
    const header = getColumnTraitByName(
      "spreadsheetConfig",
      "tableMenuSpace",
      "header",
    );
    const colIndex = liveColIndex(meta, {
      header,
      columnId: getColumnTraitByName(
        "spreadsheetConfig",
        "tableMenuSpace",
        "columnId",
      ),
    });
    if (colIndex === undefined || colIndex === meta.fullTableColIndexes[0]) {
      return;
    }
    throw new Error(
      `${floorColumnLabel(header)} is no longer the first column of ${this.sheet.raw.title}'s Table. Undo the move, or move it back to the Table's first column.`,
    );
  }
  createMissing(): string[] {
    const meta = this.sheet.raw.meta;
    return this._missingColumns().map((floorColumn) => {
      const colIndex = meta.nextEndColumnInsertIndex;
      meta.addSheetChangeToSave({
        action: "insertColumn",
        startColumnIndex: colIndex,
      });
      meta
        .column(colIndex)
        .updateUniformCell("tableHeader", floorColumn.header)
        .updateUniformCell("columnId", floorColumn.columnId)
        .updateUniformCell("colGroupName", floorColumn.groupHeading);
      return `${this.sheet.raw.title} · ${floorColumn.header} (${floorColumn.columnId})`;
    });
  }
  private _missingColumns(): FloorColumnRestore[] {
    const meta = this.sheet.raw.meta;
    return floorColumnsToRestore(this.sheetName).filter(
      (floorColumn) => liveColIndex(meta, floorColumn) === undefined,
    );
  }
}

export function floorRecreatableColumns<SN extends FloorSheetName>(
  sheetName: SN,
): readonly ColumnName<SN>[] {
  return recreatableColumns()[sheetName];
}

// Only columns the sync or an endpoint refills by itself; recreating any other empty loses what it declared.
function recreatableColumns(): {
  [SN in FloorSheetName]: readonly ColumnName<SN>[];
} {
  return {
    spreadsheetConfig: spreadsheetConfigFeedbackColumnNames(),
    sheetConfig: ["sheetTitle"],
    columnConfig: ["sheetTitle", "header"],
  };
}

function floorColumnLabel(header: string): string {
  return `Floor column "${header}"`;
}
