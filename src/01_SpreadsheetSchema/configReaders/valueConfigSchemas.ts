import { Val, validationError } from "@byronbroughten/utils/val";

import { frameworkValueSchemas } from "../../00_Source/CellValues/frameworkValueSchemas";
import {
  type ValueSchemaBase,
  vsc,
} from "../../00_Source/CellValues/valueSchema";
import type { FixedList, FixedListMember } from "./makeConfigs";
import {
  getValueConfig,
  isFixedList,
  type ValueConfigName,
  valueConfigNames,
  type ValueConfigValue,
} from "./valueConfigsTypes";

type ValueConfigSchemas = {
  [K in ValueConfigName]: ValueSchemaBase<ValueConfigValue<K>>;
};

export function makeSchemasFromValueConfig(): ValueConfigSchemas {
  return valueConfigNames().reduce((schemas, name) => {
    (schemas[name] as ValueSchemaBase<ValueConfigValue<typeof name>>) =
      makeValueConfigSchema(name) as ValueSchemaBase<
        ValueConfigValue<typeof name>
      >;
    return schemas;
  }, {} as ValueConfigSchemas);
}

// A live reference reads and writes as a string; never the id schema, whose default throws.
function makeValueConfigSchema(name: ValueConfigName): ValueSchemaBase {
  const valueConfig = getValueConfig(name);
  if (isFixedList(valueConfig)) {
    return makeFixedListSchema(name, valueConfig);
  } else {
    return frameworkValueSchemas.string;
  }
}

function makeFixedListSchema(
  name: ValueConfigName,
  fixedList: FixedList,
): ValueSchemaBase {
  return vsc<FixedListMember>({
    type: fixedList[0] as FixedListMember,
    makeDefault: () => Val.assert(fixedList[0], `'${name}' first member`),
    strictValidate: (value: unknown) => {
      if (fixedList.includes(value as FixedListMember)) {
        return value as FixedListMember;
      } else {
        throw validationError(value, `'${name}' union value element.`);
      }
    },
    blankReadsAs: null,
  });
}
