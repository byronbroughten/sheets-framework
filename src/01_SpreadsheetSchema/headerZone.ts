import {
  SheetIndex,
  type SheetRowIndex,
} from "../00_Source/RawSource/SheetIndex";
import { headRows } from "./headRows";
import { tableLayout } from "./tableLayout";

const bounds = {
  depth: tableLayout.headerZoneDepth,
  // The highest header row whose head rows still fit on the sheet.
  topHeaderRowIndex: headRows.countAboveHeader,
} as const;

function holdsHeaderRow(sheetRowIndex: SheetRowIndex): boolean {
  return (
    sheetRowIndex >= bounds.topHeaderRowIndex && sheetRowIndex < bounds.depth
  );
}

function headerRowsLabel(): string {
  const top = bounds.topHeaderRowIndex + 1;
  if (top === bounds.depth) return `row ${top}`;
  return `rows ${top}–${bounds.depth}`;
}

// The sheet's top rows, across every column, where each managed Table keeps its header row.
export const headerZone = {
  endRowIndex: SheetIndex.row(bounds.depth),
  headerRowsLabel: headerRowsLabel(),
  holdsHeaderRow,
  holdsActionRow(sheetRowIndex: SheetRowIndex): boolean {
    return holdsHeaderRow(
      SheetIndex.row(sheetRowIndex + tableLayout.headRowOffsets.action),
    );
  },
};
