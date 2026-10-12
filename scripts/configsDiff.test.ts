import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { describe, expect, it } from "vitest";

import {
  type ConfigSnapshot,
  readConfigsDiff,
  summarizeConfigsDiff,
} from "./configsDiff.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";

const item = { tableId: "item", tableName: "item", hasIdColumn: true };
const itemName = { columnId: "c:itm:aaa", header: "Name", valueName: "string" };
const itemQty = { columnId: "c:itm:bbb", header: "Qty", valueName: "number" };

function snapshot(overrides: Partial<ConfigSnapshot> = {}): ConfigSnapshot {
  return {
    tableConfigs: { item },
    columnConfigs: { item: { name: itemName, qty: itemQty } },
    valueConfigs: { status: ["Open", "Closed"] },
    ...overrides,
  };
}

describe("summarizeConfigsDiff", () => {
  it("says no changes when nothing changed", () => {
    expect(
      summarizeConfigsDiff({ before: snapshot(), after: snapshot() }),
    ).toEqual(["No config changes."]);
  });

  it("lists added and removed Tables by name", () => {
    const log = { tableId: "log", tableName: "log", hasIdColumn: false };
    const after = snapshot({
      tableConfigs: { log },
      columnConfigs: { log: {} },
    });
    expect(summarizeConfigsDiff({ before: snapshot(), after })).toEqual([
      "Table item: removed",
      "  column name: removed",
      "  column qty: removed",
      "Table log: added",
      "4 changes.",
    ]);
  });

  it("reports a Table with the same tableId under a new key as one rename", () => {
    const after = snapshot({
      tableConfigs: { goods: item },
      columnConfigs: { goods: { name: itemName, qty: itemQty } },
    });
    expect(summarizeConfigsDiff({ before: snapshot(), after })).toEqual([
      "Table item → goods: renamed",
      "1 change.",
    ]);
  });

  it("lists added, removed and renamed columns under their Table", () => {
    const itemNote = {
      columnId: "c:itm:ccc",
      header: "Note",
      valueName: "string",
    };
    const after = snapshot({
      columnConfigs: { item: { title: itemName, note: itemNote } },
    });
    expect(summarizeConfigsDiff({ before: snapshot(), after })).toEqual([
      "Table item:",
      "  column name → title: renamed",
      "  column qty: removed",
      "  column note: added",
      "3 changes.",
    ]);
  });

  it("reports changed fields on a Table and on a column with old and new values", () => {
    const after = snapshot({
      tableConfigs: { item: { ...item, hasIdColumn: false } },
      columnConfigs: {
        item: { name: itemName, qty: { ...itemQty, valueName: "string" } },
      },
    });
    expect(summarizeConfigsDiff({ before: snapshot(), after })).toEqual([
      "Table item: hasIdColumn true → false",
      '  column qty: valueName "number" → "string"',
      "2 changes.",
    ]);
  });

  it("reports value configs added, removed and changed", () => {
    const before = snapshot({
      valueConfigs: { status: ["Open", "Closed"], size: ["S"] },
    });
    const after = snapshot({
      valueConfigs: { status: ["Open", "Pending"], color: ["Red"] },
    });
    expect(summarizeConfigsDiff({ before, after })).toEqual([
      'Value status: added "Pending"; removed "Closed"',
      "Value size: removed",
      "Value color: added",
      "3 changes.",
    ]);
  });

  it("diffs value configs by kind: fixed list or live reference", () => {
    const liveReference = { kind: "liveReference" } as const;
    const before = snapshot({
      valueConfigs: {
        status: ["Open", "Closed"],
        size: ["S", "M"],
        agent: ["Ann"],
        tenant: liveReference,
        priority: liveReference,
      },
    });
    const after = snapshot({
      valueConfigs: {
        status: liveReference,
        size: ["S", 2],
        agent: ["Ann"],
        tenant: liveReference,
        priority: [1, 2],
      },
    });
    expect(summarizeConfigsDiff({ before, after })).toEqual([
      "Value status: fixed list → live reference",
      'Value size: added 2; removed "M"',
      "Value priority: live reference → fixed list",
      "3 changes.",
    ]);
  });

  it("tells a number member from the same digits as a string", () => {
    const before = snapshot({ valueConfigs: { size: ["1", 2] } });
    const after = snapshot({ valueConfigs: { size: [1, 2] } });
    expect(summarizeConfigsDiff({ before, after })).toEqual([
      'Value size: added 1; removed "1"',
      "1 change.",
    ]);
  });

  it("counts everything as added on a first generation with no HEAD version", () => {
    expect(
      summarizeConfigsDiff({ before: undefined, after: snapshot() }),
    ).toEqual([
      "Table item: added",
      "  column name: added",
      "  column qty: added",
      "Value status: added",
      "4 changes.",
    ]);
  });
});

describe("readConfigsDiff", () => {
  function writeConfigFiles(generatedDir: string, valueConfigs: object): void {
    const files = { tableConfigs: {}, columnConfigs: {}, valueConfigs };
    Object.entries(files).forEach(([file, configs]) => {
      const maker = `make${file[0]?.toUpperCase()}${file.slice(1)}`;
      writeFileSync(
        join(generatedDir, `${file}.ts`),
        `export const ${file} = ${maker}(${JSON.stringify(configs, null, 2)});\n`,
      );
    });
  }

  function git(cwd: string, ...args: string[]): void {
    const config = ["-c", "user.name=test", "-c", "user.email=test@test"];
    const { status, stderr } = spawnSync("git", [...config, ...args], {
      cwd,
      encoding: "utf8",
    });
    if (status !== 0) throw new Error(stderr);
  }

  it("reads live references and number lists from the generated files", () => {
    const dir = mkdtempSync(join(tmpdir(), "configs-diff-"));
    const generatedDir = join(dir, "generated");
    mkdirSync(generatedDir);
    writeConfigFiles(generatedDir, { status: ["Open"], size: ["S"] });
    git(dir, "init", "-q");
    git(dir, "add", ".");
    git(dir, "commit", "-q", "-m", "HEAD");
    writeConfigFiles(generatedDir, {
      status: { kind: "liveReference" },
      size: [1, 2],
    });
    const sheetsConfig: SheetsConfig = {
      path: join(dir, "sheets.config.json"),
      dir,
      spreadsheetId: "",
      generatedDir,
      choreHomes: [],
      endpointModule: undefined,
    };
    expect(readConfigsDiff(sheetsConfig)).toEqual([
      "Value status: fixed list → live reference",
      'Value size: added 1, 2; removed "S"',
      "2 changes.",
    ]);
  });
});
