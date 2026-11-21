import { Val } from "@byronbroughten/utils/val";

import { emptyStateRaw } from "../ClassTypes/emptyStateRaw.ts";
import type { SheetStateRaw } from "../ClassTypes/StateRaw.ts";
import {
  SpreadsheetBaseRaw,
  type SpreadsheetRawProps,
} from "./SpreadsheetBaseRaw.ts";

export interface SheetRawProps extends SpreadsheetRawProps {
  sheetGid: number;
}

export class SheetBaseRaw extends SpreadsheetBaseRaw {
  declare readonly sheetGid: number;
  constructor({ sheetGid, ...props }: SheetRawProps) {
    super(props);
    this.sheetGid = sheetGid;
    this._assertSheetNotQueuedForDelete();
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
  // A write reaching the tab would be sent before its deleteSheet, then lost with it.
  private _assertSheetNotQueuedForDelete(): void {
    if (!this.sheetIsQueuedForDelete(this.sheetGid)) return;
    throw new Error(
      `Sheet GID ${this.sheetGid} is queued for deletion in this flush; nothing more may be queued on it or its Tables.`,
    );
  }
  protected _ensureSheetState(): void {
    if (this.sheetsStateRaw.has(this.sheetGid)) return;
    this.sheetsStateRaw.set(this.sheetGid, emptyStateRaw.sheetState());
  }
}
