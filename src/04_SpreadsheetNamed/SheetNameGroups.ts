import {
  type SheetConfigs,
  sheetConfigsByName,
} from "../01_SpreadsheetSchema/sheetConfigsTypes";
import { type SubType } from "../utils/Obj";

export type SheetNameWithIdColumn = keyof SubType<
  SheetConfigs,
  { hasIdColumn: true }
>;

export type SheetNameWithNameColumn = keyof SubType<
  SheetConfigs,
  { hasNameColumn: true }
>;

export type SheetNameWithIdAndNameColumn = SheetNameWithIdColumn &
  SheetNameWithNameColumn;

interface SheetNameGroups {
  hasIdColumn: SheetNameWithIdColumn[];
}
export type TnGroupName = keyof SheetNameGroups;

export type SheetNameByGroup<GN extends TnGroupName> =
  SheetNameGroups[GN][number];

export function isInTnGroup<GN extends TnGroupName>(
  groupName: GN,
  sn: string,
): sn is SheetNameByGroup<GN> {
  return sheetConfigsByName()[sn]?.[groupName] === true;
}
