import { lazy } from "../utils/lazy";
import { Obj } from "../utils/Obj";
import { Val } from "../utils/Val";
import { type Configs, installedConfigs } from "./configRegister";
import type { SheetConfigsBase, SheetConfigStored } from "./makeConfigs";

// Post-sheetConfigs
export type SheetConfigs = Configs["sheetConfigs"];
export type SheetNameSimple = keyof SheetConfigs & string;
export function configSheetNames(): SheetNameSimple[] {
  return Obj.keys(sheetConfigs()) as SheetNameSimple[];
}
export type SheetName<TN extends SheetNameSimple = SheetNameSimple> = TN;
export interface SheetConfig<
  HI extends boolean = boolean,
> extends SheetConfigStored<HI> {
  sheetName: string;
}

export function getSheetTraitByName<
  TN extends SheetNameSimple,
  TK extends keyof SheetConfig,
>(sheetName: TN, key: TK): SheetConfig[TK] {
  if (key === "sheetName") {
    return sheetName as unknown as SheetConfig[TK];
  }
  return sheetConfigs()[sheetName][
    key as keyof SheetConfigStored
  ] as SheetConfig[TK];
}

export const sheetConfigsByGid = lazy(() =>
  Obj.toKeyedMap(sheetConfigs(), "sheetGid", "sheetName"),
);

export function getSheetTraitByGid<TK extends keyof SheetConfig>(
  sheetGid: number,
  key: TK,
): SheetConfig[TK] {
  return Val.assert(
    sheetConfigsByGid().get(sheetGid),
    `Sheet with gid ${sheetGid}`,
  )[key];
}

// The generated literal read by an arbitrary name, where an entry may be absent.
export function sheetConfigsByName(): SheetConfigsBase {
  return sheetConfigs();
}

function sheetConfigs(): SheetConfigs {
  return installedConfigs().sheetConfigs;
}
