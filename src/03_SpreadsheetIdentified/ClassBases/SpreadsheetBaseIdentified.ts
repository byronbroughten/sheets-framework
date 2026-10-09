import {
  SpreadsheetBaseRaw,
  type SpreadsheetRawProps,
} from "../../02_SpreadsheetRaw/ClassBases/SpreadsheetBaseRaw";
import type { StateIdentified } from "../ClassTypes/StateIdentified";
import {
  type FeedbackColumnIds,
  installedFeedbackColumnIds,
} from "../feedbackColumnRegister";

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
    feedbackColumnIds: FeedbackColumnIds = installedFeedbackColumnIds(),
  ): SpreadsheetIdentifiedProps {
    return {
      ...SpreadsheetBaseRaw.initSpreadsheetRawProps(),
      spreadsheetStateIdentified: { tables: new Map() },
      feedbackColumnIds,
    };
  }
  protected get tablesStateIdentified(): StateIdentified["tables"] {
    return this.spreadsheetStateIdentified.tables;
  }
}
