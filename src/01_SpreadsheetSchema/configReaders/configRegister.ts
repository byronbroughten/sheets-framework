import type {
  IdNameMapGeneric,
  IdNameMapTab,
  IdNameMapTable,
} from "./idNameMapTypes";
import type {
  ColumnConfigsBase,
  TabConfigsBase,
  TableConfigsBase,
  ValueConfigsBase,
} from "./makeConfigs";

// The tab list and retired prefixes are optional so a configs set generated before them still registers.
export interface ConfigSetBase {
  tableConfigs: TableConfigsBase;
  tabConfigs?: TabConfigsBase;
  retiredIdPrefixes?: readonly string[];
  columnConfigs: ColumnConfigsBase;
  valueConfigs: ValueConfigsBase;
}

// Empty until the program augments it once with `configs: typeof appConfigs`.
// eslint-disable-next-line @typescript-eslint/no-empty-object-type -- declaration merging fills it
export interface Register {}

export interface ConfigsNotRegistered {
  "Augment Register with `configs: typeof appConfigs`": never;
}

export type ConfigsOf<RG> = RG extends {
  configs: infer CS extends ConfigSetBase;
}
  ? CS
  : ConfigsNotRegistered;

export type Configs = ConfigsOf<Register>;

// Names nothing, so an app that hasn't generated its map yet still type-checks an empty `alterations.ts`.
interface IdNameMapEmpty {
  tables: Record<never, IdNameMapTable>;
  tabs: Record<never, IdNameMapTab>;
}

export type IdNameMapOf<RG> = RG extends {
  idNameMap: infer IM extends IdNameMapGeneric;
}
  ? IM
  : IdNameMapEmpty;

export type IdNameMapRegistered = IdNameMapOf<Register>;

let installed: Configs | undefined;

// The lazy derivations cache the first set, so a second would be silently ignored.
export function installConfigs(configs: Configs): void {
  if (installed !== undefined && installed !== configs) {
    throw new Error(
      "A different set of configs is already installed. A program installs one set, once.",
    );
  }
  installed = configs;
}

export function installedConfigs(): Configs {
  if (installed === undefined) {
    throw new Error(
      "Configs have not been installed. The entry call must supply the app's configs before anything reads them.",
    );
  }
  return installed;
}
