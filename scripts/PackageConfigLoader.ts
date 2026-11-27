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
  static fromSheetsConfig({ generatedDir }: SheetsConfig): PackageConfigLoader {
    return new PackageConfigLoader(generatedDir);
  }

  private readonly generatedDir: string;
  constructor(generatedDir: string) {
    this.generatedDir = generatedDir;
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
      tableConfigs: await this._importConfig("tableConfigs"),
      tabConfigs: await this._importConfigIfPresent("tabConfigs"),
      conditionalFormatConfigs: await this._importConfigIfPresent(
        "conditionalFormatConfigs",
      ),
      retiredIdPrefixes: await this._importConfigIfPresent("retiredIdPrefixes"),
      columnConfigs: await this._importConfig("columnConfigs"),
      valueConfigs: await this._importConfig("valueConfigs"),
    };
    // The registered set may or may not carry the optional files, and these are the generated files it registers.
    return configs as Configs;
  }

  private async _importConfig<CF extends ConfigFile>(
    base: CF,
  ): Promise<ConfigSetBase[CF]> {
    const url = pathToFileURL(this.filePath(base));
    return (await import(url.href))[base];
  }

  // A file the package's last regen predates is absent, so its config reads as unset.
  private async _importConfigIfPresent<CF extends ConfigFile>(
    base: CF,
  ): Promise<ConfigSetBase[CF] | undefined> {
    return importExportIfPresent<ConfigSetBase[CF]>(this.filePath(base), base);
  }
}
