// `sheets-framework gen-configs`: regenerates the package's generated files from its live Tables and previous configs, on the Node host.
import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";

import type { ConfigRegeneration } from "../src/05_Operators/SyncCoordinator.ts";
import { readConfigsDiff } from "./configsDiff.ts";
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
  readonly path: Record<ConfigFile | "appConfigs" | "idNameMap", string>;
  constructor({ sheetsConfig }: { sheetsConfig: SheetsConfig }) {
    this.sheetsConfig = sheetsConfig;
    const { generatedDir } = sheetsConfig;
    this.path = {
      tableConfigs: configFilePath(generatedDir, "tableConfigs"),
      tabConfigs: configFilePath(generatedDir, "tabConfigs"),
      conditionalFormatConfigs: configFilePath(
        generatedDir,
        "conditionalFormatConfigs",
      ),
      retiredIdPrefixes: configFilePath(generatedDir, "retiredIdPrefixes"),
      columnConfigs: configFilePath(generatedDir, "columnConfigs"),
      valueConfigs: configFilePath(generatedDir, "valueConfigs"),
      appConfigs: join(generatedDir, "appConfigs.ts"),
      idNameMap: idNameMapPath(generatedDir),
    };
  }
  static init(sheetsConfig: SheetsConfig): ConfigFilesGenerator {
    return new ConfigFilesGenerator({ sheetsConfig });
  }
  async run(): Promise<void> {
    const {
      tableConfigs,
      tabConfigs,
      conditionalFormatConfigs,
      retiredIdPrefixes,
      columnConfigs,
      valueConfigs,
      idNameMap,
      untypedColumnsSummary,
    } = await this._generate();

    // Write nothing until every config and the ID-name map are confirmed good; a subset would go stale.
    mkdirSync(this.sheetsConfig.generatedDir, { recursive: true });
    writeFileSync(this.path.tableConfigs, tableConfigs);
    writeFileSync(this.path.tabConfigs, tabConfigs);
    writeFileSync(this.path.conditionalFormatConfigs, conditionalFormatConfigs);
    writeFileSync(this.path.retiredIdPrefixes, retiredIdPrefixes);
    writeFileSync(this.path.columnConfigs, columnConfigs);
    writeFileSync(this.path.valueConfigs, valueConfigs);
    writeFileSync(this.path.appConfigs, appConfigsText());
    writeFileSync(this.path.idNameMap, idNameMap);
    Object.values(this.path).forEach((path) => {
      console.log(`Wrote ${path}`);
    });
    this._createAlterationsIfMissing();
    console.log("\ngen:configs: changes against HEAD:");
    console.log(readConfigsDiff(this.sheetsConfig).join("\n"));
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

  _createAlterationsIfMissing(): void {
    const path = alterationsPath(this.sheetsConfig.generatedDir);
    if (existsSync(path)) {
      return;
    }
    writeFileSync(path, alterationsText());
    console.log(`Created ${path}`);
  }

  async _generate(): Promise<ConfigRegeneration> {
    // The previous configs say which Tables have API access; a package with no generated files yet borrows the framework's.
    // No app endpoints: they may name columns the configs being regenerated don't have yet.
    await startNodeHost({
      isDryRun: false,
      sheetsConfig: this.sheetsConfig,
      configs: hasPackageConfigs(this.sheetsConfig)
        ? await loadPackageConfigs(this.sheetsConfig)
        : await loadFrameworkConfigs(),
    });
    const { SyncCoordinator } =
      await import("../src/05_Operators/SyncCoordinator.ts");
    return SyncCoordinator.init().generateConfigFiles();
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
    "\ngen:configs: regeneration succeeded and every generated file was written," +
      "but this package's `npm run tsc` failed above. This usually means " +
      "hand-written references in this package still name a sheet or column " +
      "that no longer exists after this regeneration. Fix those references " +
      "and re-run `npm run tsc` — do not hand-edit the generated files.",
  );
}

export function idNameMapPath(generatedDir: string): string {
  return join(generatedDir, "idNameMap.ts");
}

// A fixed path from the generated directory, so finding it needs no setting of its own.
export function alterationsPath(generatedDir: string): string {
  return join(dirname(generatedDir), "alterations.ts");
}

export function alterationsText(): string {
  return `import { makePendingAlterations } from "@byronbroughten/sheets-framework";

export const pendingAlterations = makePendingAlterations({});
`;
}

export function appConfigsText(): string {
  return `// Without it, a program that reaches this file only by dynamic import rejects the augmentation (TS2664).
import type {} from "@byronbroughten/sheets-framework";

import { columnConfigs } from "./columnConfigs";
import { conditionalFormatConfigs } from "./conditionalFormatConfigs";
import type { IdNameMap } from "./idNameMap";
import { retiredIdPrefixes } from "./retiredIdPrefixes";
import { tabConfigs } from "./tabConfigs";
import { tableConfigs } from "./tableConfigs";
import { valueConfigs } from "./valueConfigs";

export const appConfigs = {
  tableConfigs,
  tabConfigs,
  conditionalFormatConfigs,
  retiredIdPrefixes,
  columnConfigs,
  valueConfigs,
};

declare module "@byronbroughten/sheets-framework" {
  interface Register {
    configs: typeof appConfigs;
    idNameMap: IdNameMap;
  }
}
`;
}

export async function runGenConfigs(sheetsConfig: SheetsConfig): Promise<void> {
  await ConfigFilesGenerator.init(sheetsConfig).run();
}
