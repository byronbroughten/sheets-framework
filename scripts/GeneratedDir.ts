import { existsSync, readFileSync } from "node:fs";
import { dirname, join } from "node:path";

import type { IdNameMapGeneric } from "../src/01_SpreadsheetSchema/configReaders/idNameMapTypes.ts";
import type { PendingAlterations } from "../src/01_SpreadsheetSchema/configReaders/intendedState.ts";
import { alterationsFileSource } from "../src/05_Operators/configFileSource.ts";
import { parseIdNameMapFileSource } from "../src/05_Operators/parseIdNameMapFileSource.ts";
import type { SyncIntendedState } from "../src/05_Operators/SyncCoordinator.ts";
import {
  type ConfigFile,
  configFilePath,
  importExportIfPresent,
} from "./nodeHost.ts";

export class GeneratedDir {
  readonly path: Record<
    ConfigFile | "appConfigs" | "idNameMap" | "alterations",
    string
  >;
  constructor(dir: string) {
    this.path = {
      tableConfigs: configFilePath(dir, "tableConfigs"),
      tabConfigs: configFilePath(dir, "tabConfigs"),
      conditionalFormatConfigs: configFilePath(dir, "conditionalFormatConfigs"),
      retiredIdPrefixes: configFilePath(dir, "retiredIdPrefixes"),
      columnConfigs: configFilePath(dir, "columnConfigs"),
      valueConfigs: configFilePath(dir, "valueConfigs"),
      appConfigs: join(dir, "appConfigs.ts"),
      idNameMap: join(dir, "idNameMap.ts"),
      // A fixed path from the generated directory, so finding it needs no setting of its own.
      alterations: join(dirname(dir), "alterations.ts"),
    };
  }
  static get alterationsText(): string {
    return alterationsFileSource({});
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
