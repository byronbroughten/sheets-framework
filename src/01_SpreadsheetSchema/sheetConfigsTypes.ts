import { lazy } from "../utils/lazy";
import { Obj } from "../utils/Obj";
import { type Configs, installedConfigs } from "./configRegister";
import type { SheetConfigsBase, SheetConfigStored } from "./makeConfigs";

// Post-sheetConfigs
export type SheetConfigs = Configs["sheetConfigs"];
type SheetNameSimple = keyof SheetConfigs & string;
export interface SheetConfig<
  HI extends boolean = boolean,
> extends SheetConfigStored<HI> {
  sheetName: string;
}

export function getSheetTraitByName<
  SN extends SheetNameSimple,
  TK extends keyof SheetConfig,
>(sheetName: SN, key: TK): SheetConfig[TK] {
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

// The generated literal read by an arbitrary name, where an entry may be absent.
export function sheetConfigsByName(): SheetConfigsBase {
  return sheetConfigs();
}

function sheetConfigs(): SheetConfigs {
  return installedConfigs().sheetConfigs;
}
