import { existsSync, readFileSync } from "node:fs";
import { dirname, join } from "node:path";

import type { IdNameMapGeneric } from "../src/01_SpreadsheetSchema/configReaders/idNameMapTypes.ts";
import type { PendingAlterations } from "../src/01_SpreadsheetSchema/configReaders/intendedState.ts";
import { alterationsFileSource } from "../src/05_Operators/configFileSource.ts";
import { parseIdNameMapFileSource } from "../src/05_Operators/parseIdNameMapFileSource.ts";
import type { SyncIntendedState } from "../src/05_Operators/SyncCoordinator.ts";
import { type ConfigFile, PackageConfigLoader } from "./PackageConfigLoader.ts";
import { importExportIfPresent } from "./packageLoader.ts";

export class GeneratedDir {
  static readonly alterationsText = alterationsFileSource({});
  static readonly appConfigsText = `// Without it, a program that reaches this file only by dynamic import rejects the augmentation (TS2664).
import type {} from "@byronbroughten/sheets-framework";

import { columnConfigs } from "./columnConfigs.ts";
import { conditionalFormatConfigs } from "./conditionalFormatConfigs.ts";
import type { IdNameMap } from "./idNameMap.ts";
import { retiredIdPrefixes } from "./retiredIdPrefixes.ts";
import { tabConfigs } from "./tabConfigs.ts";
import { tableConfigs } from "./tableConfigs.ts";
import { valueConfigs } from "./valueConfigs.ts";

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
  readonly dir: string;
  readonly path: Record<
    ConfigFile | "appConfigs" | "idNameMap" | "alterations",
    string
  >;
  constructor(dir: string) {
    this.dir = dir;
    const configLoader = new PackageConfigLoader(dir);
    this.path = {
      tableConfigs: configLoader.filePath("tableConfigs"),
      tabConfigs: configLoader.filePath("tabConfigs"),
      conditionalFormatConfigs: configLoader.filePath(
        "conditionalFormatConfigs",
      ),
      retiredIdPrefixes: configLoader.filePath("retiredIdPrefixes"),
      columnConfigs: configLoader.filePath("columnConfigs"),
      valueConfigs: configLoader.filePath("valueConfigs"),
      appConfigs: join(dir, "appConfigs.ts"),
      idNameMap: join(dir, "idNameMap.ts"),
      // A fixed path from the generated directory, so finding it needs no setting of its own.
      alterations: join(dirname(dir), "alterations.ts"),
    };
  }

  // Intended state is the pending alterations alone until claims arrive; a package with no alterations.ts has none pending.
  async loadIntendedState(): Promise<SyncIntendedState> {
    const pendingAlterations = await importExportIfPresent<PendingAlterations>(
      this.path.alterations,
      "pendingAlterations",
    );
    return { alterations: pendingAlterations ?? {} };
  }

  // A package that has never synced has no map, so every Table is new to it.
  loadPreviousIdNameMap(): IdNameMapGeneric {
    const { idNameMap } = this.path;
    if (!existsSync(idNameMap)) return { tables: {}, tabs: {} };
    return parseIdNameMapFileSource(readFileSync(idNameMap, "utf8"));
  }
}
