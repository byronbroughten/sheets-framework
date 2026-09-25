import { installedRawSource } from "../../00_Source/RawSource/RawSource";
import { SpreadsheetBaseSchema } from "../../01_SpreadsheetSchema/SpreadsheetBaseSchema";
import type { GridRangeProps } from "../ClassTypes/AccessorsRaw";
import { emptyStateRaw } from "../ClassTypes/emptyStateRaw";
import type { SheetsStateRaw, StateRaw } from "../ClassTypes/StateRaw";

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
  get schema(): SpreadsheetBaseSchema {
    return new SpreadsheetBaseSchema();
  }
  get fetcherGridRanges(): GridRangeProps[] {
    return this.spreadsheetStateRaw.fetchQueue.gridRanges;
  }
  get updateRequests(): StateRaw["writeQueue"]["updateRequests"] {
    return this.spreadsheetStateRaw.writeQueue.updateRequests;
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
        timeZone: undefined,
        rawSource: installedRawSource(),
        fetchQueue: emptyStateRaw.spreadsheetFetchQueue(),
        writeQueue: emptyStateRaw.spreadsheetWriteQueue(),
        sheets: new Map(),
      },
    };
  }
}
