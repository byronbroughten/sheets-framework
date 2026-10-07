import type { SheetChange } from "../../00_Source/PlatformEvents/sheetChange";
import {
  configSheetFloorSeed,
  type FloorTabName,
} from "../../01_SpreadsheetSchema/configSheetFloorSeed";
import { getTableTraitByName } from "../../01_SpreadsheetSchema/tableConfigsTypes";

export interface FloorNotice {
  title: string;
  message: string;
  untilClosed: boolean;
}

const unwarnedFloorTab: FloorTabName = "valueConfig";

export function floorChangeNotice(
  change: SheetChange,
  liveTitlesByGid: ReadonlyMap<number, string>,
): FloorNotice | undefined {
  const { title } = configSheetFloorSeed[unwarnedFloorTab];
  const liveTitle = liveTitlesByGid.get(
    getTableTraitByName(unwarnedFloorTab, "sheetGid"),
  );
  if (change === "other" && liveTitle !== undefined && liveTitle !== title) {
    return {
      title: `${title} is managed`,
      message: `This tab keeps the name "${title}". Your rename will switch back the next time configs sync.`,
      untilClosed: false,
    };
  }
  if (change === "sheetRemoved" && liveTitle === undefined) {
    return {
      title: `${title} was deleted`,
      message:
        "Press Undo (Ctrl+Z, or ⌘Z on a Mac) now to get it back with its data. If you don't, the next sync recreates it empty.",
      untilClosed: true,
    };
  }
  return undefined;
}
