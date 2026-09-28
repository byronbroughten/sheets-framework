import {
  SpreadsheetBaseIdentified,
  type SpreadsheetIdentifiedProps,
} from "../../03_SpreadsheetIdentified/ClassBases/SpreadsheetBaseIdentified";
import type { FeedbackColumnIds } from "../../03_SpreadsheetIdentified/feedbackColumnRegister";

export interface SpreadsheetNamedProps extends SpreadsheetIdentifiedProps {}

export class SpreadsheetBaseNamed extends SpreadsheetBaseIdentified {
  get spreadsheetNamedProps(): SpreadsheetNamedProps {
    return {
      ...this.spreadsheetIdentifiedProps,
    };
  }
  static initSpreadsheetNamedProps(
    feedbackColumnIds?: FeedbackColumnIds,
  ): SpreadsheetNamedProps {
    return SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(
      feedbackColumnIds,
    );
  }
}
