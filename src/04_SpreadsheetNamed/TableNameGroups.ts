import {
  type TableConfigs,
  tableConfigsByName,
} from "../01_SpreadsheetSchema/tableConfigsTypes";
import { type SubType } from "../utils/Obj";

export type SheetNameWithIdColumn = keyof SubType<
  TableConfigs,
  { hasIdColumn: true }
>;

export type SheetNameWithNameColumn = keyof SubType<
  TableConfigs,
  { hasNameColumn: true }
>;

export type SheetNameWithIdAndNameColumn = SheetNameWithIdColumn &
  SheetNameWithNameColumn;

interface TableNameGroups {
  hasIdColumn: SheetNameWithIdColumn[];
}
export type TnGroupName = keyof TableNameGroups;

export type SheetNameByGroup<GN extends TnGroupName> =
  TableNameGroups[GN][number];

export function isInTnGroup<GN extends TnGroupName>(
  groupName: GN,
  sn: string,
): sn is SheetNameByGroup<GN> {
  return tableConfigsByName()[sn]?.[groupName] === true;
}
