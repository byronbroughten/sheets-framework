import {
  SpreadsheetBaseRaw,
  type SpreadsheetRawProps,
} from "../../02_SpreadsheetRaw/ClassBases/SpreadsheetBaseRaw";
import type { StateIdentified } from "../ClassTypes/StateIdentified";

// Column ids by sheet GID; handed down as data, since the endpoints that declare them sit tiers above.
export type FeedbackColumnIds = ReadonlyMap<number, ReadonlySet<string>>;

export interface SpreadsheetIdentifiedProps extends SpreadsheetRawProps {
  spreadsheetStateIdentified: StateIdentified;
  feedbackColumnIds: FeedbackColumnIds;
}

export class SpreadsheetBaseIdentified extends SpreadsheetBaseRaw {
  protected spreadsheetStateIdentified: StateIdentified;
  readonly feedbackColumnIds: FeedbackColumnIds;
  constructor({
    spreadsheetStateIdentified,
    feedbackColumnIds,
    ...rest
  }: SpreadsheetIdentifiedProps) {
    super(rest);
    this.spreadsheetStateIdentified = spreadsheetStateIdentified;
    this.feedbackColumnIds = feedbackColumnIds;
  }
  get spreadsheetIdentifiedProps(): SpreadsheetIdentifiedProps {
    return {
      ...this.spreadsheetRawProps,
      spreadsheetStateIdentified: this.spreadsheetStateIdentified,
      feedbackColumnIds: this.feedbackColumnIds,
    };
  }
  static initSpreadsheetIdentifiedProps(
    feedbackColumnIds: FeedbackColumnIds = new Map(),
  ): SpreadsheetIdentifiedProps {
    return {
      ...SpreadsheetBaseRaw.initSpreadsheetRawProps(),
      spreadsheetStateIdentified: {
        sheets: new Map(),
      },
      feedbackColumnIds,
    };
  }
  get sheetsStateIdentified(): StateIdentified["sheets"] {
    return this.spreadsheetStateIdentified.sheets;
  }
}
