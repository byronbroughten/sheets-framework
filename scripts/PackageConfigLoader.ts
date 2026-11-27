import { existsSync } from "node:fs";
import { join } from "node:path";
import { pathToFileURL } from "node:url";

import type {
  Configs,
  ConfigSetBase,
} from "../src/01_SpreadsheetSchema/configReaders/configRegister.ts";
import { importExportIfPresent } from "./packageLoader.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";

// An app's configs register without the rest, so only these decide whether it has configs of its own.
const requiredConfigFiles = [
  "tableConfigs",
  "columnConfigs",
  "valueConfigs",
] as const;

export type ConfigFile =
  | (typeof requiredConfigFiles)[number]
  | "tabConfigs"
  | "conditionalFormatConfigs"
  | "retiredIdPrefixes";

export class PackageConfigLoader {
  readonly generatedDir: string;
  constructor(generatedDir: string) {
    this.generatedDir = generatedDir;
  }

  static fromSheetsConfig({ generatedDir }: SheetsConfig): PackageConfigLoader {
    return new PackageConfigLoader(generatedDir);
  }

  filePath(base: ConfigFile): string {
    return join(this.generatedDir, `${base}.ts`);
  }

  hasConfigs(): boolean {
    return requiredConfigFiles.every((base) => existsSync(this.filePath(base)));
  }

  // The generated files in the package's generatedDir, as the package's entry passes them to the framework.
  async loadConfigs(): Promise<Configs> {
    const configs: ConfigSetBase = {
      tableConfigs: await this.importConfig("tableConfigs"),
      tabConfigs: await this.importConfigIfPresent("tabConfigs"),
      conditionalFormatConfigs: await this.importConfigIfPresent(
        "conditionalFormatConfigs",
      ),
      retiredIdPrefixes: await this.importConfigIfPresent("retiredIdPrefixes"),
      columnConfigs: await this.importConfig("columnConfigs"),
      valueConfigs: await this.importConfig("valueConfigs"),
    };
    // The registered set may or may not carry the optional files, and these are the generated files it registers.
    return configs as Configs;
  }

  async importConfig<CF extends ConfigFile>(
    base: CF,
  ): Promise<ConfigSetBase[CF]> {
    const url = pathToFileURL(this.filePath(base));
    return (await import(url.href))[base];
  }

  // A file the package's last regen predates is absent, so its config reads as unset.
  async importConfigIfPresent<CF extends ConfigFile>(
    base: CF,
  ): Promise<ConfigSetBase[CF] | undefined> {
    return importExportIfPresent<ConfigSetBase[CF]>(this.filePath(base), base);
  }
}
