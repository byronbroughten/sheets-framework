import {
  configSheetFloorSeed,
  type FloorSeedColumn,
  floorSeedColumns,
  type FloorTabName,
} from "../../01_SpreadsheetSchema/configSheetFloorSeed";
import { dimensionIds } from "../../01_SpreadsheetSchema/dimensionIds";
import { getTableTraitByName } from "../../01_SpreadsheetSchema/tableConfigsTypes";
import { TableOrigin } from "../../01_SpreadsheetSchema/TableOrigin";
import { SpreadsheetBaseNamed } from "../../04_SpreadsheetNamed/ClassBases/SpreadsheetBaseNamed";
import { SpreadsheetNamed } from "../../04_SpreadsheetNamed/SpreadsheetNamed";
import { type FloorSheetName, floorSheetNames } from "./floorSeedLookups";
import { FloorTabColumnCreator } from "./FloorTabColumnCreator";

const creatableFloorTabNames = [
  "spreadsheetConfig",
  "tableConfig",
  "columnConfig",
  "valueConfig",
] as const satisfies readonly FloorTabName[];

const exampleColumn = configSheetFloorSeed.valueConfig.exampleColumn;

type CreatedTableColumn = Pick<FloorSeedColumn, "header" | "columnType">;

/**
 * Creates a missing floor tab at its generated GID, with its seeded Table at
 * its generated tableId and placed by the sheet layout, and recreates missing
 * floor columns at their Table's end, with the generated column ID, seeded
 * header and group heading, failing closed on a missing column the sync can't
 * refill. ConfigSheetFloor runs this right after its fetch and flushes only
 * when it reports something. Each tab's recreatable table and insert live in
 * FloorTabColumnCreator.
 * docs/generated-data/config-sheet-floor.md
 */
export class ConfigSheetFloorCreator extends SpreadsheetBaseNamed {
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  createMissing(): string[] {
    const report: string[] = [];
    const createdTabs = this._createMissingTabs();
    if (createdTabs.length > 0) {
      report.push(`Created tabs: ${createdTabs.join("; ")}`);
    }
    const tabs = floorSheetNames()
      .map((sheetName) => this._floorTab(sheetName))
      .filter((tab) => tab.hasFloorTable());
    tabs.forEach((tab) => tab.assertMissingAreRecreatable());
    const createdLines = tabs.flatMap((tab) => tab.createMissing());
    if (createdLines.length > 0) {
      report.push(`Recreated columns: ${createdLines.join("; ")}`);
    }
    return report;
  }
  private _createMissingTabs(): string[] {
    const origin = TableOrigin.expected();
    return creatableFloorTabNames.flatMap((sheetName) => {
      const sheetGid = getTableTraitByName(sheetName, "sheetGid");
      if (this.ss.raw.gidIsActive(sheetGid)) return [];
      const seed = configSheetFloorSeed[sheetName];
      const columns = createdTableColumns(sheetName);
      const endRowIdx = origin.sheetRowIndex(createdDataRowCount(sheetName));
      const endColIdx = origin.sheetColIndex(columns.length);
      this.ss.raw
        .gatherAddSheetOperation({
          sheetId: sheetGid,
          title: seed.title,
          rowCount: endRowIdx,
          columnCount: endColIdx,
        })
        .gatherAddTableOperation({
          tableId: getTableTraitByName(sheetName, "tableId"),
          name: seed.tableName,
          range: {
            sheetId: sheetGid,
            startRowIndex: origin.headerRowIndex,
            endRowIndex: endRowIdx,
            startColumnIndex: origin.startColIndex,
            endColumnIndex: endColIdx,
          },
          columnProperties: columns.map((column, columnIndex) => ({
            columnIndex,
            columnName: column.header,
            columnType: column.columnType,
          })),
        });
      if (sheetName === "valueConfig") {
        this._seedExampleColumn(sheetGid, origin);
      }
      return [seed.title];
    });
  }
  // The add-Table's columnName writes the header, so no header cell is written.
  private _seedExampleColumn(sheetGid: number, origin: TableOrigin): void {
    const colIndex = origin.sheetColIndex(0);
    const idPrefix = getTableTraitByName("valueConfig", "idPrefix");
    this.ss.raw.gatherAddedSheetFillCellOperation({
      sheetId: sheetGid,
      rowIndex: origin.headSheetRowIndex("columnId"),
      colIndex,
      value: dimensionIds.col(idPrefix),
    });
    exampleColumn.seededValues.forEach((value, memberIndex) => {
      this.ss.raw.gatherAddedSheetFillCellOperation({
        sheetId: sheetGid,
        rowIndex: origin.sheetRowIndex(memberIndex),
        colIndex,
        value,
      });
    });
  }
  private _floorTab(
    sheetName: FloorSheetName,
  ): FloorTabColumnCreator<FloorSheetName> {
    return new FloorTabColumnCreator({
      ...this.spreadsheetNamedProps,
      tableName: sheetName,
    });
  }
}

function createdTableColumns(
  sheetName: FloorTabName,
): readonly CreatedTableColumn[] {
  if (sheetName === "valueConfig") return [exampleColumn];
  return floorSeedColumns(sheetName);
}

function createdDataRowCount(sheetName: FloorTabName): number {
  if (sheetName === "valueConfig") return exampleColumn.seededValues.length;
  return 1;
}
