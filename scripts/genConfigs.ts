// `sheets-framework gen-configs`: regenerates the package's four config files from its live config sheets, on the Node host.
import { spawnSync } from "node:child_process";
import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, relative } from "node:path";
import { fileURLToPath } from "node:url";

import type { ConfigRegeneration } from "../src/05_Operators/ConfigCoordinator.ts";
import {
  type ConfigFile,
  configFilePath,
  hasPackageConfigs,
  loadFrameworkConfigs,
  loadPackageConfigs,
  startNodeHost,
} from "./nodeHost.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";

class ConfigFilesGenerator {
  readonly sheetsConfig: SheetsConfig;
  readonly path: Record<ConfigFile, string>;
  constructor({ sheetsConfig }: { sheetsConfig: SheetsConfig }) {
    this.sheetsConfig = sheetsConfig;
    const { generatedDir } = sheetsConfig;
    this.path = {
      spreadsheetConfig: configFilePath(generatedDir, "spreadsheetConfig"),
      sheetConfigs: configFilePath(generatedDir, "sheetConfigs"),
      columnConfigs: configFilePath(generatedDir, "columnConfigs"),
      valueConfigs: configFilePath(generatedDir, "valueConfigs"),
    };
  }
  static init(sheetsConfig: SheetsConfig): ConfigFilesGenerator {
    return new ConfigFilesGenerator({ sheetsConfig });
  }
  async run(): Promise<void> {
    const {
      spreadsheetConfig,
      sheetConfigs,
      columnConfigs,
      valueConfigs,
      untypedColumnsSummary,
      floorReport,
      idPrefixReport,
      declaredCellReport,
    } = await this._generate();

    // Write nothing until all four are confirmed good; a subset would go stale.
    mkdirSync(dirname(this.path.spreadsheetConfig), { recursive: true });
    writeFileSync(this.path.spreadsheetConfig, spreadsheetConfig);
    writeFileSync(this.path.sheetConfigs, sheetConfigs);
    writeFileSync(this.path.columnConfigs, columnConfigs);
    writeFileSync(this.path.valueConfigs, valueConfigs);
    console.log(`Wrote ${this.path.spreadsheetConfig}`);
    console.log(`Wrote ${this.path.sheetConfigs}`);
    console.log(`Wrote ${this.path.columnConfigs}`);
    console.log(`Wrote ${this.path.valueConfigs}`);
    if (floorReport !== "") {
      console.log(`\ngen:configs: ${floorReport}`);
    }
    if (idPrefixReport !== undefined) {
      console.log(`\ngen:configs: ${idPrefixReport}`);
    }
    if (declaredCellReport !== undefined) {
      console.log(`\ngen:configs: ${declaredCellReport}`);
    }
    console.log(
      `\ngen:configs: ${untypedColumnsSummary ?? "every column is declared; no value name was guessed."}`,
    );

    console.log(
      "\nRunning this package's npm run tsc to check the regenerated files...",
    );
    if (!this._runTsc()) {
      reportTscFailure();
      process.exit(1);
    }
    console.log("gen:configs: tsc passed.");
  }

  async _generate(): Promise<ConfigRegeneration> {
    // Only the config floor is read here; a package with no generated files yet borrows the framework's.
    await startNodeHost({
      isDryRun: false,
      sheetsConfig: this.sheetsConfig,
      configs: hasPackageConfigs(this.sheetsConfig)
        ? await loadPackageConfigs(this.sheetsConfig)
        : await loadFrameworkConfigs(),
    });
    const { ConfigCoordinator } =
      await import("../src/05_Operators/ConfigCoordinator.ts");
    return ConfigCoordinator.init().generateConfigFiles(
      this._makeConfigsImport(),
    );
  }

  _makeConfigsImport(): string {
    const makeConfigsPath = fileURLToPath(
      new URL("../src/01_SpreadsheetSchema/makeConfigs", import.meta.url),
    );
    return relative(dirname(this.path.spreadsheetConfig), makeConfigsPath);
  }

  _runTsc(): boolean {
    const { status } = spawnSync("npm", ["run", "tsc"], {
      cwd: this.sheetsConfig.dir,
      stdio: "inherit",
    });
    return status === 0;
  }
}

function reportTscFailure(): void {
  console.error(
    "\ngen:configs: regeneration succeeded and all four files were written, " +
      "but this package's `npm run tsc` failed above. This usually means " +
      "hand-written references in this package still name a sheet or column " +
      "that no longer exists after this regeneration. Fix those references " +
      "and re-run `npm run tsc` — do not hand-edit the generated files.",
  );
}

export async function runGenConfigs(sheetsConfig: SheetsConfig): Promise<void> {
  await ConfigFilesGenerator.init(sheetsConfig).run();
}
