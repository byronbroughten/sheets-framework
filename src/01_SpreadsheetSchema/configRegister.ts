import type {
  ColumnConfigsBase,
  SheetConfigsBase,
  SpreadsheetConfigBase,
  ValueConfigsBase,
} from "./makeConfigs";

export interface ConfigSetBase {
  spreadsheetConfig: SpreadsheetConfigBase;
  sheetConfigs: SheetConfigsBase;
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
