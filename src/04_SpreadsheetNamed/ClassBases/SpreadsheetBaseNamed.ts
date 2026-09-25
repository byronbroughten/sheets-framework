import {
  SpreadsheetBaseIdentified,
  type SpreadsheetIdentifiedProps,
} from "../../03_SpreadsheetIdentified/ClassBases/SpreadsheetBaseIdentified";

export interface SpreadsheetNamedProps extends SpreadsheetIdentifiedProps {}

export class SpreadsheetBaseNamed extends SpreadsheetBaseIdentified {
  get spreadsheetNamedProps(): SpreadsheetNamedProps {
    return {
      ...this.spreadsheetIdentifiedProps,
    };
  }
  static initSpreadsheetNamedProps(): SpreadsheetNamedProps {
    return SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps();
  }
}
