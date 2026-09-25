import type { GridRangeProps as BaseGridRangeProps } from "../../00_Source/RawSource/RawSource";
import type { StrictOmit, StrictPick } from "../../utils/Obj";

export type GridRangeProps = BaseGridRangeProps;

export type SheetGridRangeProps = StrictOmit<GridRangeProps, "sheetId">;
export type ColumnGridRangeProps = StrictOmit<
  GridRangeProps,
  "sheetId" | "startColumnIndex" | "endColumnIndex"
>;

export type SheetColumnsRange = StrictPick<
  GridRangeProps,
  "startColumnIndex" | "endColumnIndex" | "sheetId"
>;
export type ColumnRange = StrictPick<
  GridRangeProps,
  "startColumnIndex" | "endColumnIndex"
>;
