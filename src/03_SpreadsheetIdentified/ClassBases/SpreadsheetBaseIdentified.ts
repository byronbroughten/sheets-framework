import {
  SpreadsheetBaseRaw,
  type SpreadsheetRawProps,
} from "../../02_SpreadsheetRaw/ClassBases/SpreadsheetBaseRaw";
import type { StateIdentified } from "../ClassTypes/StateIdentified";

export interface SpreadsheetIdentifiedProps extends SpreadsheetRawProps {
  spreadsheetStateIdentified: StateIdentified;
}

export class SpreadsheetBaseIdentified extends SpreadsheetBaseRaw {
  protected spreadsheetStateIdentified: StateIdentified;
  constructor({
    spreadsheetStateIdentified,
    ...rest
  }: SpreadsheetIdentifiedProps) {
    super(rest);
    this.spreadsheetStateIdentified = spreadsheetStateIdentified;
  }
  get spreadsheetIdentifiedProps(): SpreadsheetIdentifiedProps {
    return {
      ...this.spreadsheetRawProps,
      spreadsheetStateIdentified: this.spreadsheetStateIdentified,
    };
  }
  static initSpreadsheetIdentifiedProps(): SpreadsheetIdentifiedProps {
    return {
      ...SpreadsheetBaseRaw.initSpreadsheetRawProps(),
      spreadsheetStateIdentified: {
        sheets: new Map(),
      },
    };
  }
  get sheetsStateIdentified(): StateIdentified["sheets"] {
    return this.spreadsheetStateIdentified.sheets;
  }
}
