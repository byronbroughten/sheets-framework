import { installedRawSource } from "../../00_Source/RawSource/RawSource";
import { SpreadsheetBaseSchema } from "../../01_SpreadsheetSchema/SpreadsheetBaseSchema";
import type { GridRangeProps } from "../ClassTypes/AccessorsRaw";
import { emptyStateRaw } from "../ClassTypes/emptyStateRaw";
import type {
  SheetsStateRaw,
  StateRaw,
  TablesStateRaw,
  WriteOperations,
} from "../ClassTypes/StateRaw";

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
  get spreadsheetRawProps(): SpreadsheetRawProps {
    return {
      spreadsheetStateRaw: this.spreadsheetStateRaw,
    };
  }
  static initSpreadsheetRawProps(): SpreadsheetRawProps {
    return {
      spreadsheetStateRaw: {
        allSheetPropertiesAreFetched: false,
        isRegeneratingConfigs: false,
        timeZone: undefined,
        rawSource: installedRawSource(),
        fetchQueue: emptyStateRaw.spreadsheetFetchQueue(),
        writeQueue: emptyStateRaw.spreadsheetWriteQueue(),
        sheets: new Map(),
        tables: new Map(),
      },
    };
  }
}
