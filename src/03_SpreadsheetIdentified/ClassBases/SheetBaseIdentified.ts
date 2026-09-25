import { Val } from "../../utils/Val";
import {
  emptySheetFetchQueueIdentified,
  type FetchTargetIdentified,
  type SheetStateIdentified,
} from "../ClassTypes/StateIdentified";
import {
  SpreadsheetBaseIdentified,
  type SpreadsheetIdentifiedProps,
} from "./SpreadsheetBaseIdentified";

export interface SheetIdentifiedProps extends SpreadsheetIdentifiedProps {
  sheetGid: number;
}
export class SheetBaseIdentified extends SpreadsheetBaseIdentified {
  readonly sheetGid: number;
  constructor(props: SheetIdentifiedProps) {
    super(props);
    this.sheetGid = props.sheetGid;
    this._ensureSheetState();
  }
  get sheetIdentifiedProps(): SheetIdentifiedProps {
    return {
      ...this.spreadsheetIdentifiedProps,
      sheetGid: this.sheetGid,
    };
  }
  private _ensureSheetState(): void {
    if (!this.sheetsStateIdentified.has(this.sheetGid)) {
      this.sheetsStateIdentified.set(this.sheetGid, {
        fetchQueue: emptySheetFetchQueueIdentified(),
      });
    }
  }
  protected get sheetState(): SheetStateIdentified {
    return Val.assert(
      this.sheetsStateIdentified.get(this.sheetGid),
      `sheetState for sheetGid ${this.sheetGid}`,
    );
  }
  get fetchTargets(): FetchTargetIdentified[] {
    return this.sheetState.fetchQueue.targets;
  }
  get isPreppedToFetch(): boolean {
    return (
      this.fetchTargets.length > 0 ||
      this.sheetState.fetchQueue.gatherConditionalFormats ||
      this.sheetState.fetchQueue.gatherEditProtections
    );
  }
  clearFetchTargets(): void {
    this.sheetState.fetchQueue = emptySheetFetchQueueIdentified();
  }
}
