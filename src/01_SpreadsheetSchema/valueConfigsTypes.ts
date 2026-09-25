import { Obj } from "../utils/Obj";
import { type Configs, installedConfigs } from "./configRegister";

export type ValueConfigs = Configs["valueConfigs"];
export type ValueConfigName = keyof ValueConfigs;
export function valueConfigNames(): readonly ValueConfigName[] {
  return Obj.keys(installedConfigs().valueConfigs);
}

export type ValueConfigValues = {
  [K in ValueConfigName]: ValueConfigs[K][number];
};
export type ValueConfigValue<VC extends ValueConfigName = ValueConfigName> =
  ValueConfigValues[VC];

export function getValueConfigValueArr<VC extends ValueConfigName>(
  key: VC,
): ValueConfigs[VC] {
  return installedConfigs().valueConfigs[key];
}
