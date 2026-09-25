import type { CellValue } from "../../00_Source/CellValues/cellValues";
import {
  configSheetFloorSeed,
  type FloorSeedColumn,
  floorSeedColumns,
  type FloorTabName,
} from "../../01_SpreadsheetSchema/configSheetFloorSeed";
import { dimensionIds } from "../../01_SpreadsheetSchema/dimensionIds";
import { getSheetTraitByName } from "../../01_SpreadsheetSchema/sheetConfigsTypes";
import {
  spreadsheetConfigIndexHeaders,
  spreadsheetConfigTextHeaders,
} from "../../01_SpreadsheetSchema/spreadsheetConfigFields";
import { ssConfigGet } from "../../01_SpreadsheetSchema/spreadsheetConfigTypes";
import { SpreadsheetBaseNamed } from "../../04_SpreadsheetNamed/ClassBases/SpreadsheetBaseNamed";
import { SpreadsheetNamed } from "../../04_SpreadsheetNamed/SpreadsheetNamed";
import { Obj } from "../../utils/Obj";
import { type FloorSheetName, floorSheetNames } from "./floorSeedLookups";
import { FloorTabColumnCreator } from "./FloorTabColumnCreator";

const creatableFloorTabNames = [
  "spreadsheetConfig",
  "sheetConfig",
  "columnConfig",
  "valueConfig",
] as const satisfies readonly FloorTabName[];

const exampleColumn = configSheetFloorSeed.valueConfig.exampleColumn;

type CreatedTableColumn = Pick<FloorSeedColumn, "header" | "columnType">;

/**
 * Creates a missing floor tab at its generated GID, with its seeded Table placed
 * by the generated layout, and recreates missing floor columns at their Table's
 * end, with the generated column ID, seeded header and group heading, failing
 * closed on a missing column the sync can't refill. ConfigSheetFloor runs this
 * right after its fetch and flushes only when it reports something. Each tab's
 * recreatable table and insert live in FloorTabColumnCreator.
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
    const headerRowIdx = ssConfigGet("tableHeaderRowIndexBase0");
    const startColIdx = ssConfigGet("startTableColIndexBase0");
    return creatableFloorTabNames.flatMap((sheetName) => {
      const sheetGid = getSheetTraitByName(sheetName, "sheetGid");
      if (this.ss.raw.gidIsActive(sheetGid)) return [];
      const seed = configSheetFloorSeed[sheetName];
      const columns = createdTableColumns(sheetName);
      const endRowIdx = headerRowIdx + 1 + createdDataRowCount(sheetName);
      const endColIdx = startColIdx + columns.length;
      this.ss.raw
        .gatherAddSheetRequest({
          sheetId: sheetGid,
          title: seed.title,
          rowCount: endRowIdx,
          columnCount: endColIdx,
        })
        .gatherAddTableRequest({
          name: seed.tableName,
          range: {
            sheetId: sheetGid,
            startRowIndex: headerRowIdx,
            endRowIndex: endRowIdx,
            startColumnIndex: startColIdx,
            endColumnIndex: endColIdx,
          },
          columnProperties: columns.map((column, columnIndex) => ({
            columnIndex,
            columnName: column.header,
            columnType: column.columnType,
          })),
        });
      if (sheetName === "spreadsheetConfig") {
        this._seedLayoutValues(sheetGid, columns);
      }
      if (sheetName === "valueConfig") {
        this._seedExampleColumn(sheetGid, headerRowIdx + 1);
      }
      return [seed.title];
    });
  }
  private _seedLayoutValues(
    sheetGid: number,
    columns: readonly CreatedTableColumn[],
  ): void {
    const headers = columns.map((column) => column.header);
    seededLayoutValues()
      .map(({ header, value }) => {
        const columnIndex = headers.indexOf(header);
        if (columnIndex === -1) {
          throw new Error(
            `Spreadsheet Config seed has no column "${header}" to seed its layout value into.`,
          );
        }
        return { columnIndex, value };
      })
      .sort((a, b) => a.columnIndex - b.columnIndex)
      .forEach(({ columnIndex, value }) => {
        this.ss.raw.gatherAddedSheetCellRequest({
          sheetId: sheetGid,
          rowIndex: ssConfigGet("tableHeaderRowIndexBase0") + 1,
          colIndex: ssConfigGet("startTableColIndexBase0") + columnIndex,
          value,
        });
      });
  }
  // The add-Table's columnName writes the header, so no header cell is written.
  private _seedExampleColumn(sheetGid: number, topDataRowIdx: number): void {
    const colIndex = ssConfigGet("startTableColIndexBase0");
    const idPrefix = getSheetTraitByName("valueConfig", "idPrefix");
    this.ss.raw.gatherAddedSheetCellRequest({
      sheetId: sheetGid,
      rowIndex: ssConfigGet("columnIdRowIdxBase0"),
      colIndex,
      value: dimensionIds.col(idPrefix),
    });
    exampleColumn.seededValues.forEach((value, memberIndex) => {
      this.ss.raw.gatherAddedSheetCellRequest({
        sheetId: sheetGid,
        rowIndex: topDataRowIdx + memberIndex,
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
      sheetName,
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

// Read through ssConfigGet, the same layout that placed the created Table.
function seededLayoutValues(): { header: string; value: CellValue }[] {
  return [
    ...Obj.keys(spreadsheetConfigTextHeaders).map((key) => ({
      header: spreadsheetConfigTextHeaders[key],
      value: ssConfigGet(key),
    })),
    ...Obj.keys(spreadsheetConfigIndexHeaders).map((key) => ({
      header: spreadsheetConfigIndexHeaders[key],
      value: ssConfigGet(key) + 1,
    })),
  ];
}
