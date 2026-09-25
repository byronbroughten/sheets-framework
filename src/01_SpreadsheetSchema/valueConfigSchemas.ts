import { type ValueSchemaBase, vsc } from "../00_Source/CellValues/valueSchema";
import { validationError } from "../utils/Val";
import {
  getValueConfigValueArr,
  type ValueConfigName,
  valueConfigNames,
  type ValueConfigValue,
} from "./valueConfigsTypes";

function makeDefaultValueConfigValue<VN extends ValueConfigName>(
  valueName: VN,
): ValueConfigValue<VN> {
  return getValueConfigValueArr(valueName)[0] as ValueConfigValue<VN>;
}
function validateValueConfigValue<VC extends ValueConfigName>(
  value: unknown,
  valueName: VC,
): ValueConfigValue<VC> {
  if (
    (getValueConfigValueArr(valueName) as readonly unknown[]).includes(value)
  ) {
    return value as ValueConfigValue<VC>;
  } else {
    throw validationError(value, `'${valueName}' union value element.`);
  }
}
type ValueConfigSchemas = {
  [K in ValueConfigName]: ValueSchemaBase<ValueConfigValue<K>>;
};

export function makeSchemasFromValueConfig(): ValueConfigSchemas {
  return valueConfigNames().reduce((schemas, name) => {
    (schemas[name] as ValueSchemaBase<ValueConfigValue<typeof name>>) = vsc({
      type: makeDefaultValueConfigValue(name) as ValueConfigValue<typeof name>,
      makeDefault: () => makeDefaultValueConfigValue(name),
      strictValidate: (value: unknown) => validateValueConfigValue(value, name),
      blankReadsAs: null,
    }) as ValueSchemaBase<ValueConfigValue<typeof name>>;
    return schemas;
  }, {} as ValueConfigSchemas);
}
