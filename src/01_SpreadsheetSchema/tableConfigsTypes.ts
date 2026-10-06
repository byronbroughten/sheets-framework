import { lazy } from "../utils/lazy";
import { Obj } from "../utils/Obj";
import { type Configs, installedConfigs } from "./configRegister";

export type TableConfigs = Configs["tableConfigs"];

export const tableConfigsByTableId = lazy(() =>
  Obj.toKeyedMap(installedConfigs().tableConfigs, "tableId", "tableKey"),
);
