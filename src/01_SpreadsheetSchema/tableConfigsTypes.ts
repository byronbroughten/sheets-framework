import { lazy } from "../utils/lazy";
import { Obj } from "../utils/Obj";
import { Val } from "../utils/Val";
import { type Configs, installedConfigs } from "./configRegister";
import type { TableConfigsBase, TableConfigStored } from "./makeConfigs";

export type TableConfigs = Configs["tableConfigs"];
export type TableNameSimple = keyof TableConfigs & string;
export function configTableNames(): TableNameSimple[] {
  return Obj.keys(tableConfigs()) as TableNameSimple[];
}
export type TableName<TN extends TableNameSimple = TableNameSimple> = TN;
export interface TableConfig extends TableConfigStored {
  tableKey: string;
}

export function getTableTraitByName<
  TN extends TableNameSimple,
  TK extends keyof TableConfig,
>(tableKey: TN, key: TK): TableConfig[TK] {
  if (key === "tableKey") {
    return tableKey as unknown as TableConfig[TK];
  }
  return tableConfigs()[tableKey][
    key as keyof TableConfigStored
  ] as TableConfig[TK];
}

export const tableConfigsByTableId = lazy(() =>
  Obj.toKeyedMap(tableConfigs(), "tableId", "tableKey"),
);

// A sheet holds one Table for now, so a GID still names one entry.
export const tableConfigsByGid = lazy(() =>
  Obj.toKeyedMap(tableConfigs(), "sheetGid", "tableKey"),
);

export const tableKeysByGid = lazy(() =>
  configTableNames().reduce((byGid, tableKey) => {
    const sheetGid = getTableTraitByName(tableKey, "sheetGid");
    byGid.set(sheetGid, [...(byGid.get(sheetGid) ?? []), tableKey]);
    return byGid;
  }, new Map<number, TableNameSimple[]>()),
);

export function getTableTraitByGid<TK extends keyof TableConfig>(
  sheetGid: number,
  key: TK,
): TableConfig[TK] {
  return Val.assert(
    tableConfigsByGid().get(sheetGid),
    `Table config for sheet gid ${sheetGid}`,
  )[key];
}

// The generated literal read by an arbitrary name, where an entry may be absent.
export function tableConfigsByName(): TableConfigsBase {
  return tableConfigs();
}

function tableConfigs(): TableConfigs {
  return installedConfigs().tableConfigs;
}
