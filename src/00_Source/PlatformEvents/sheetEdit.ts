import type { SheetColIndex, SheetRowIndex } from "../RawSource/SheetIndex";

export interface SheetEdit {
  sheetGid: number;
  rowIndexBase0: SheetRowIndex;
  colIndexBase0: SheetColIndex;
  value: string | undefined;
}
