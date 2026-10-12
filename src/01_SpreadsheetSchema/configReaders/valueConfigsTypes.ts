import { Obj } from "@byronbroughten/utils/obj";

import { type Configs, installedConfigs } from "./configRegister";
import type { FixedList, ValueConfigBase } from "./makeConfigs";

export type ValueConfigs = Configs["valueConfigs"];
export type ValueConfigName = keyof ValueConfigs;
export function valueConfigNames(): readonly ValueConfigName[] {
  return Obj.keys(installedConfigs().valueConfigs);
}

// A live reference's members are rows added after the build, so it is only ever a string.
type ValueConfigMember<VB extends ValueConfigBase> = VB extends FixedList
  ? VB[number]
  : string;
export type ValueConfigValues = {
  [K in ValueConfigName]: ValueConfigMember<ValueConfigs[K]>;
};
export type ValueConfigValue<VC extends ValueConfigName = ValueConfigName> =
  ValueConfigValues[VC];

export function getValueConfig<VC extends ValueConfigName>(
  key: VC,
): ValueConfigs[VC] {
  return installedConfigs().valueConfigs[key];
}

export function isFixedList(
  valueConfig: ValueConfigBase,
): valueConfig is FixedList {
  return Array.isArray(valueConfig);
}
