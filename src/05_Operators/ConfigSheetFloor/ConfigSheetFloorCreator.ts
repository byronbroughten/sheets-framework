import {
  configSheetFloorSeed,
  floorSeedColumns,
  type FloorTabName,
} from "../../01_SpreadsheetSchema/configReaders/configSheetFloorSeed";
import { getTableTraitByName } from "../../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { TableOrigin } from "../../01_SpreadsheetSchema/TableOrigin";
import { SpreadsheetBaseNamed } from "../../04_SpreadsheetNamed/ClassBases/SpreadsheetBaseNamed";
import { SpreadsheetNamed } from "../../04_SpreadsheetNamed/SpreadsheetNamed";
import { floorTabNames } from "./floorSeedLookups";
import { FloorTabColumnCreator } from "./FloorTabColumnCreator";

// A created floor Table holds one empty data row.
const createdDataRowCount = 1;

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
    const tabs = floorTabNames()
      .map((tableName) => this._floorTab(tableName))
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
    return floorTabNames().flatMap((tableName) => {
      const sheetGid = getTableTraitByName(tableName, "sheetGid");
      if (this.ss.raw.gidIsActive(sheetGid)) return [];
      const seed = configSheetFloorSeed[tableName];
      const columns = floorSeedColumns(tableName);
      const endRowIdx = origin.sheetRowIndex(createdDataRowCount);
      const endColIdx = origin.sheetColIndex(columns.length);
      this.ss.raw
        .gatherAddSheetOperation({
          sheetId: sheetGid,
          title: seed.title,
          rowCount: endRowIdx,
          columnCount: endColIdx,
        })
        .gatherAddTableOperation({
          tableId: getTableTraitByName(tableName, "tableId"),
          name: seed.liveTableName,
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
      return [seed.title];
    });
  }
  private _floorTab(
    tableName: FloorTabName,
  ): FloorTabColumnCreator<FloorTabName> {
    return new FloorTabColumnCreator({
      ...this.spreadsheetNamedProps,
      tableName: tableName,
    });
  }
}
