import { existsSync } from "node:fs";
import { basename } from "node:path";
import { pathToFileURL } from "node:url";

import type { DropdownStyleReset } from "../src/00_Source/RawSource/RawSource.ts";
import type { Configs } from "../src/01_SpreadsheetSchema/configReaders/configRegister.ts";
import type { Endpoints } from "../src/06_API/Endpoints.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";

// One named export of a module that may not exist yet; undefined when the file is absent.
export async function importExportIfPresent<MX>(
  modulePath: string,
  exportName: string,
): Promise<MX | undefined> {
  if (!existsSync(modulePath)) return undefined;
  return (await import(pathToFileURL(modulePath).href))[exportName];
}

export async function loadPackageEndpoints({
  endpointModule,
}: SheetsConfig): Promise<Endpoints | undefined> {
  if (endpointModule === undefined) return undefined;
  return importConstNamedAfterFile<Endpoints>(
    endpointModule,
    "An endpoint module",
  );
}

// The endpoint module's dropdownStyleReset export; undefined when the app has no module or the module exports none.
export async function loadDropdownStyleReset({
  endpointModule,
}: SheetsConfig): Promise<DropdownStyleReset | undefined> {
  if (endpointModule === undefined) return undefined;
  const setting = await importExportIfPresent<unknown>(
    endpointModule,
    "dropdownStyleReset",
  );
  if (setting === undefined) return undefined;
  if (setting === "ask" || setting === "allow") return setting;
  throw new Error(
    `${endpointModule} exports a dropdownStyleReset of ${JSON.stringify(setting)}; it is "ask" or "allow".`,
  );
}

// The one rule a chore file and an endpoint module share, so a typo names the file and the const it lacks.
export async function importConstNamedAfterFile<MX>(
  modulePath: string,
  fileKind: string,
): Promise<MX> {
  const name = basename(modulePath).replace(/\.ts$/, "");
  const value = (await import(pathToFileURL(modulePath).href))[name];
  if (!value) {
    throw new Error(
      `${modulePath} exports no "${name}". ${fileKind} exports one const named after the file.`,
    );
  }
  return value;
}

// The framework's own dev configs: the previous configs a package's first sync reads with, before it has generated files of its own.
export async function loadFrameworkConfigs(): Promise<Configs> {
  return (await import("../dev/generated/appConfigs.ts")).appConfigs;
}
