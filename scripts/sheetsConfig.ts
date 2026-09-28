// Finds a package's sheets.config.json by walking up from cwd, the only source of its spreadsheet ID. See docs/how-it-runs.md, "The sheets-framework bin".
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";

const sheetsConfigFileNames = {
  config: "sheets.config.json",
  example: "sheets.config.example.json",
} as const;
// Where an app keeps each, relative to its config, unless the config names another.
const conventionalPaths = {
  endpointModule: "src/businessEndpoints.ts",
  generatedDir: "src/generated",
  choreHomes: ["src/chores", "src/chores/oneOff"],
} as const;
const siblingScan = {
  depth: 3,
  skippedDirs: new Set(["node_modules", "dist", "coverage"]),
} as const;

// A package's sheets.config.json, its folders resolved to absolute paths.
export interface SheetsConfig {
  path: string;
  dir: string;
  spreadsheetId: string;
  generatedDir: string;
  choreHomes: string[];
  // Absent when the package has no endpoint module; chores then see only the framework's endpoints.
  endpointModule: string | undefined;
}

export function loadSheetsConfig(cwd = process.cwd()): SheetsConfig {
  const path = nearestConfigPath(resolve(cwd));
  const config = readSheetsConfig(path);
  assertDistinctIds([
    path,
    ...siblingConfigPaths(path).filter((other) => other !== path),
  ]);
  return config;
}

function nearestConfigPath(cwd: string): string {
  for (let dir = cwd; ; dir = dirname(dir)) {
    const path = join(dir, sheetsConfigFileNames.config);
    if (existsSync(path)) return path;
    const examplePath = join(dir, sheetsConfigFileNames.example);
    if (existsSync(examplePath)) {
      throw new Error(
        `${examplePath} exists but ${path} does not. Copy the example to ${sheetsConfigFileNames.config} and fill in your spreadsheet ID.`,
      );
    }
    if (dirname(dir) === dir) {
      throw new Error(
        `No ${sheetsConfigFileNames.config} in ${cwd} or any folder above it. Run the bin from inside a package.`,
      );
    }
  }
}

function readSheetsConfig(path: string): SheetsConfig {
  const raw = JSON.parse(readFileSync(path, "utf8"));
  const dir = dirname(path);
  if (!isFilledString(raw.spreadsheetId)) {
    throw new Error(`${path} has no "spreadsheetId" string.`);
  }
  return {
    path,
    dir,
    spreadsheetId: raw.spreadsheetId,
    generatedDir: join(dir, generatedDirOf(path, raw.generatedDir)),
    choreHomes: choreHomesOf(path, raw.choreHomes).map((home) =>
      join(dir, home),
    ),
    endpointModule: endpointModuleOf(path, raw.endpointModule),
  };
}

function generatedDirOf(configPath: string, key: unknown): string {
  if (key === undefined) return conventionalPaths.generatedDir;
  if (!isFilledString(key)) {
    throw new Error(
      `${configPath} has a "generatedDir" that isn't a folder string.`,
    );
  }
  return key;
}

function choreHomesOf(configPath: string, key: unknown): readonly string[] {
  if (key === undefined) return conventionalPaths.choreHomes;
  if (!Array.isArray(key) || !key.every(isFilledString)) {
    throw new Error(
      `${configPath} has a "choreHomes" that isn't an array of folder strings.`,
    );
  }
  return key;
}

function endpointModuleOf(
  configPath: string,
  key: unknown,
): string | undefined {
  const dir = dirname(configPath);
  if (key === undefined) {
    const conventional = join(dir, conventionalPaths.endpointModule);
    return existsSync(conventional) ? conventional : undefined;
  }
  if (!isFilledString(key)) {
    throw new Error(
      `${configPath} has an "endpointModule" that isn't a file path string.`,
    );
  }
  const configured = join(dir, key);
  if (!existsSync(configured)) {
    throw new Error(
      `${configPath} names "endpointModule" ${key}, but ${configured} does not exist.`,
    );
  }
  return configured;
}

// Every package config in the repo holding this one, so a copy-pasted ID is caught from either side.
function siblingConfigPaths(configPath: string): string[] {
  const root = repoRootOf(dirname(configPath));
  const found: string[] = [];
  function visit(dir: string, depth: number): void {
    const path = join(dir, sheetsConfigFileNames.config);
    if (existsSync(path)) found.push(path);
    if (depth === siblingScan.depth) return;
    readdirSync(dir, { withFileTypes: true })
      .filter(
        (entry) =>
          entry.isDirectory() &&
          !entry.name.startsWith(".") &&
          !siblingScan.skippedDirs.has(entry.name),
      )
      .forEach((entry) => visit(join(dir, entry.name), depth + 1));
  }
  visit(root, 0);
  return found;
}

function repoRootOf(start: string): string {
  for (let dir = start; ; dir = dirname(dir)) {
    if (existsSync(join(dir, ".git"))) return dir;
    if (dirname(dir) === dir) return start;
  }
}

function assertDistinctIds(paths: string[]): void {
  const byId = new Map<string, string>();
  paths.forEach((path) => {
    const { spreadsheetId } = JSON.parse(readFileSync(path, "utf8"));
    const other = byId.get(spreadsheetId);
    if (other) {
      throw new Error(
        `${other} and ${path} share a spreadsheet ID. Fix one before running anything.`,
      );
    }
    byId.set(spreadsheetId, path);
  });
}

function isFilledString(value: unknown): value is string {
  return typeof value === "string" && value !== "";
}
