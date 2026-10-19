import { Val } from "@byronbroughten/utils/val";

import { emptyStateRaw } from "../ClassTypes/emptyStateRaw";
import type { SheetStateRaw } from "../ClassTypes/StateRaw";
import {
  SpreadsheetBaseRaw,
  type SpreadsheetRawProps,
} from "./SpreadsheetBaseRaw";

export interface SheetRawProps extends SpreadsheetRawProps {
  sheetGid: number;
}

export class SheetBaseRaw extends SpreadsheetBaseRaw {
  readonly sheetGid: number;
  constructor({ sheetGid, ...props }: SheetRawProps) {
    super(props);
    this.sheetGid = sheetGid;
  }
  get sheetRawProps(): SheetRawProps {
    return { ...this.spreadsheetRawProps, sheetGid: this.sheetGid };
  }
  get sheetState(): SheetStateRaw {
    return Val.assert(
      this.sheetsStateRaw.get(this.sheetGid),
      `sheetState for sheetGid ${this.sheetGid}`,
    );
  }
  protected _ensureSheetState(): void {
    if (this.sheetsStateRaw.has(this.sheetGid)) return;
    this.sheetsStateRaw.set(this.sheetGid, emptyStateRaw.sheetState());
  }
}
