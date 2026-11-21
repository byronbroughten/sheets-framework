import { installedRawSource } from "../../00_Source/RawSource/installedRawSource.ts";
import type { GridRangeProps } from "../../00_Source/RawSource/RawSource.ts";
import { SpreadsheetBaseSchema } from "../../01_SpreadsheetSchema/SpreadsheetBaseSchema.ts";
import { emptyStateRaw } from "../ClassTypes/emptyStateRaw.ts";
import type {
  SheetsStateRaw,
  StateRaw,
  TablesStateRaw,
  WriteOperations,
} from "../ClassTypes/StateRaw.ts";

export interface SpreadsheetRawProps {
  spreadsheetStateRaw: StateRaw;
}

export class SpreadsheetBaseRaw {
  protected spreadsheetStateRaw: StateRaw;
  constructor(props: SpreadsheetRawProps) {
    this.spreadsheetStateRaw = props.spreadsheetStateRaw;
  }
  protected get sheetsStateRaw(): SheetsStateRaw {
    return this.spreadsheetStateRaw.sheets;
  }
  protected get tablesStateRaw(): TablesStateRaw {
    return this.spreadsheetStateRaw.tables;
  }
  get schema(): SpreadsheetBaseSchema {
    return new SpreadsheetBaseSchema();
  }
  get fetcherGridRanges(): GridRangeProps[] {
    return this.spreadsheetStateRaw.fetchQueue.gridRanges;
  }
  get writeOperations(): WriteOperations {
    return this.spreadsheetStateRaw.writeQueue.operations;
  }
  protected sheetIsQueuedForDelete(sheetGid: number): boolean {
    return this.writeOperations.deleteSheet.some(
      ({ sheetId }) => sheetId === sheetGid,
    );
  }
  get spreadsheetRawProps(): SpreadsheetRawProps {
    return {
      spreadsheetStateRaw: this.spreadsheetStateRaw,
    };
  }
  protected _updateWorkingTableName(tableId: string, name: string): void {
    const properties = this.tablesStateRaw.get(tableId)?.properties;
    if (properties !== undefined) properties.name = name;
  }
  static initSpreadsheetRawProps(): SpreadsheetRawProps {
    return {
      spreadsheetStateRaw: {
        allSheetPropertiesAreFetched: false,
        usedTableIds: new Set(),
        timeZone: undefined,
        landedFlushCount: 0,
        rawSource: installedRawSource(),
        fetchQueue: emptyStateRaw.spreadsheetFetchQueue(),
        writeQueue: emptyStateRaw.spreadsheetWriteQueue(),
        sheets: new Map(),
        tables: new Map(),
      },
    };
  }
}
