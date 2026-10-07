import {
  type ColumnName,
  getColumnTraitByName,
} from "../../01_SpreadsheetSchema/columnConfigsTypes";
import { getTableTraitByName } from "../../01_SpreadsheetSchema/tableConfigsTypes";
import { TableBaseNamed } from "../../04_SpreadsheetNamed/ClassBases/TableBaseNamed";
import { SpreadsheetNamed } from "../../04_SpreadsheetNamed/SpreadsheetNamed";
import type { TableNamed } from "../../04_SpreadsheetNamed/TableNamed";
import { liveColIndex } from "./floorColumnLocation";
import {
  columnNameByHeader,
  type FloorColumnRestore,
  floorColumnsToRestore,
  type FloorSheetName,
  spreadsheetConfigFeedbackColumnNames,
} from "./floorSeedLookups";

export class FloorTabColumnCreator<
  TN extends FloorSheetName,
> extends TableBaseNamed<TN> {
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  get table(): TableNamed<TN> {
    return this.ss.table(this.tableName);
  }
  hasFloorTable(): boolean {
    const sheetGid = getTableTraitByName(this.tableName, "sheetGid");
    if (!this.ss.raw.gidIsActive(sheetGid)) return false;
    return this.table.raw.hasOneTable();
  }
  assertMissingAreRecreatable(): void {
    this._assertTableMenuSpaceIsFirst();
    const recreatable = recreatableColumns()[this.tableName];
    this._missingColumns().forEach((floorColumn) => {
      const columnName = columnNameByHeader(this.tableName, floorColumn.header);
      if (recreatable.includes(columnName)) return;
      throw new Error(
        `${floorColumnLabel(floorColumn.header)} is missing from ${this.table.raw.title}, and recreating it empty would lose what it held. Undo the delete, or insert a column headed "${floorColumn.header}" in its Table and fill it.`,
      );
    });
  }
  // Putting it back first would need a mid-Table insert, and column inserts land only at the Table end.
  private _assertTableMenuSpaceIsFirst(): void {
    if (this.tableName !== "spreadsheetConfig") return;
    const meta = this.table.raw.meta;
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
      `${floorColumnLabel(header)} is no longer the first column of ${this.table.raw.title}'s Table. Undo the move, or move it back to the Table's first column.`,
    );
  }
  createMissing(): string[] {
    return this._missingColumns().map((floorColumn) => {
      this.table.raw.appendColumn({
        columnId: floorColumn.columnId,
        header: floorColumn.header,
        groupHeading1: floorColumn.groupHeading,
      });
      return `${this.table.raw.title} · ${floorColumn.header} (${floorColumn.columnId})`;
    });
  }
  private _missingColumns(): FloorColumnRestore[] {
    const meta = this.table.raw.meta;
    return floorColumnsToRestore(this.tableName).filter(
      (floorColumn) => liveColIndex(meta, floorColumn) === undefined,
    );
  }
}

export function floorRecreatableColumns<TN extends FloorSheetName>(
  tableName: TN,
): readonly ColumnName<TN>[] {
  return recreatableColumns()[tableName];
}

// Only columns the sync or an endpoint refills by itself; recreating any other empty loses what it declared.
function recreatableColumns(): {
  [TN in FloorSheetName]: readonly ColumnName<TN>[];
} {
  return {
    spreadsheetConfig: spreadsheetConfigFeedbackColumnNames(),
    tableConfig: ["tableName", "sheetTitle"],
    columnConfig: ["tableName", "header"],
  };
}

function floorColumnLabel(header: string): string {
  return `Floor column "${header}"`;
}
