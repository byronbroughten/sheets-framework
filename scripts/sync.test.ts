import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, relative } from "node:path";
import { fileURLToPath } from "node:url";

import { Str } from "@byronbroughten/utils/str";
import {
  afterEach,
  beforeAll,
  beforeEach,
  describe,
  expect,
  it,
  vi,
} from "vitest";

import { appConfigs } from "../dev/generated/appConfigs.ts";
import type { IdNameMapGeneric } from "../src/01_SpreadsheetSchema/configReaders/idNameMapTypes.ts";
import type {
  SyncDryRunResult,
  SyncRunResult,
  SyncSendResult,
} from "../src/01_SpreadsheetSchema/configReaders/syncRunResult.ts";
import { columnIdsByName } from "../src/05_Operators/configFileSource.ts";
import { parseIdNameMapFileSource } from "../src/05_Operators/parseIdNameMapFileSource.ts";
import type {
  ConfigRegeneration,
  SyncRun,
} from "../src/05_Operators/SyncCoordinator.ts";
import type { HeldList } from "../src/nodeHost/EnteredDataGate.ts";
import type { HeldItem } from "../src/nodeHost/FlushClassifier/flushClassifierTypes.ts";
import { SyncRunReport } from "../src/nodeHost/SyncRunReport.ts";
import { GeneratedDir } from "./GeneratedDir.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";
import { parseForcedTraits, SyncRunner } from "./sync.ts";

const devIdNameMapPath = new GeneratedDir(
  fileURLToPath(new URL("../dev/generated", import.meta.url)),
).path.idNameMap;

// The diff shells out to git, which the exit code doesn't depend on.
vi.mock("./configsDiff.ts", () => ({ readConfigsDiff: () => [] }));

const runDir = mkdtempSync(join(tmpdir(), "sync-"));
const runSheetsConfig: SheetsConfig = {
  path: join(runDir, "sheets.config.json"),
  dir: runDir,
  spreadsheetId: "unused",
  generatedDir: join(runDir, "generated"),
  choreHomes: [],
  endpointModule: undefined,
};

const regeneratedFiles: ConfigRegeneration = {
  tableConfigs: "",
  tabConfigs: "",
  conditionalFormatConfigs: "",
  retiredIdPrefixes: "",
  columnConfigs: "",
  valueConfigs: "",
  idNameMap: "",
};

const dryRunResult = {
  mode: "dry run",
  target: "dev",
  command: "npm run sync",
  generatedDir: "dev/generated/",
  alterationsPath: "dev/alterations.ts",
  checks: [],
  refusals: [],
  stops: [],
  plan: [],
  alreadyMet: [],
  forcedTraits: [],
  notices: [],
  configs: {
    heldBack: [],
    renamed: [],
    namesMoved: [],
    rewrites: [],
    typedTraitChanges: [],
    notes: [],
  },
} as const satisfies Omit<SyncDryRunResult, "outcome">;

const sendResult = {
  ...dryRunResult,
  mode: "send",
  landed: [],
  notSent: [],
  didntLand: [],
  hasEmptiedAlterations: false,
} as const satisfies Omit<SyncSendResult, "outcome">;

interface ScriptRun {
  run: SyncRun;
  isTscPassing?: boolean;
}

// Runs the script past its live read: only the Node host's sync and the package's tsc are stood in for.
async function exitCodeOf({
  run,
  isTscPassing = true,
}: ScriptRun): Promise<number | string | null | undefined> {
  const runner = SyncRunner.init(runSheetsConfig, []);
  vi.spyOn(runner, "_sync").mockResolvedValue(run);
  vi.spyOn(runner, "_runTsc").mockReturnValue(isTscPassing);
  await runner.run();
  return process.exitCode;
}

describe("the sync script's exit code", () => {
  beforeEach(() => {
    vi.spyOn(console, "log").mockImplementation(() => undefined);
    vi.spyOn(console, "error").mockImplementation(() => undefined);
  });
  afterEach(() => {
    process.exitCode = undefined;
    vi.restoreAllMocks();
  });

  it("is 0 for a dry run that is ready", async () => {
    const result: SyncRunResult = {
      ...dryRunResult,
      outcome: { kind: "ready" },
    };
    expect(await exitCodeOf({ run: { result, files: regeneratedFiles } })).toBe(
      0,
    );
  });

  it("is 0 for a ready dry run with held items", async () => {
    const held = {
      block: { groups: [], fullListPath: "held.txt", token: "abc" },
      forcedGroups: [],
    };
    const result: SyncRunResult = {
      ...dryRunResult,
      outcome: { kind: "ready" },
      held,
    };
    expect(await exitCodeOf({ run: { result, files: regeneratedFiles } })).toBe(
      0,
    );
  });

  it("is 0 for a send where everything landed", async () => {
    const result: SyncRunResult = {
      ...sendResult,
      outcome: { kind: "sent", summary: "1 sent" },
    };
    expect(await exitCodeOf({ run: { result, files: regeneratedFiles } })).toBe(
      0,
    );
  });

  it("is 1 for a stopped run, dry or sent", async () => {
    const dryRun: SyncRunResult = {
      ...dryRunResult,
      outcome: { kind: "stopped" },
    };
    const send: SyncRunResult = { ...sendResult, outcome: { kind: "stopped" } };
    expect(
      await exitCodeOf({ run: { result: dryRun, files: regeneratedFiles } }),
    ).toBe(1);
    expect(
      await exitCodeOf({ run: { result: send, files: regeneratedFiles } }),
    ).toBe(1);
  });

  it("is 1 for a refused run, which writes no files", async () => {
    const result: SyncRunResult = {
      ...dryRunResult,
      outcome: { kind: "refused" },
    };
    expect(await exitCodeOf({ run: { result } })).toBe(1);
  });

  it("is 1 for a send with an item that didn't land", async () => {
    const result: SyncRunResult = {
      ...sendResult,
      outcome: { kind: "sent", summary: "0 sent" },
      didntLand: [{ subject: { table: "item" }, what: "rename" }],
    };
    expect(await exitCodeOf({ run: { result, files: regeneratedFiles } })).toBe(
      1,
    );
  });

  it("is 1 for a send a held flush stopped before any flush landed", async () => {
    const result: SyncRunResult = {
      ...sendResult,
      outcome: { kind: "stoppedBeforeSending", heldBy: "uncovered" },
    };
    expect(await exitCodeOf({ run: { result, files: regeneratedFiles } })).toBe(
      1,
    );
  });

  it("is 1 for a send a held flush stopped partway", async () => {
    const result: SyncRunResult = {
      ...sendResult,
      outcome: {
        kind: "stoppedPartway",
        landedFlushes: 1,
        flushCount: 2,
        heldBy: "mismatched",
      },
    };
    expect(await exitCodeOf({ run: { result, files: regeneratedFiles } })).toBe(
      1,
    );
  });

  it("is 1 when the package's tsc fails after a ready run", async () => {
    const result: SyncRunResult = {
      ...dryRunResult,
      outcome: { kind: "ready" },
    };
    expect(
      await exitCodeOf({
        run: { result, files: regeneratedFiles },
        isTscPassing: false,
      }),
    ).toBe(1);
  });
});

describe("the dev package's generated files", () => {
  const { tableConfigs, columnConfigs } = appConfigs;
  let idNameMap: IdNameMapGeneric;
  beforeAll(() => {
    idNameMap = parseIdNameMapFileSource(
      readFileSync(devIdNameMapPath, "utf8"),
    );
  });

  it("give each configured Table the same name, tab and column IDs in the ID-name map", () => {
    Object.entries(tableConfigs).forEach(([tableKey, tableConfig]) => {
      expect(idNameMap.tables[tableConfig.tableId]).toEqual({
        name: tableKey,
        sheetGid: tableConfig.sheetGid,
        columns: columnIdsByName(
          columnConfigs[tableKey as keyof typeof columnConfigs],
        ),
      });
    });
  });

  it("give columns in the ID-name map only to configured Tables", () => {
    const tableIdsWithColumns = Object.entries(idNameMap.tables).flatMap(
      ([tableId, table]) => (table.columns === undefined ? [] : [tableId]),
    );
    expect(tableIdsWithColumns.sort()).toEqual(
      Object.values(tableConfigs)
        .map((tableConfig) => tableConfig.tableId)
        .sort(),
    );
  });

  it("name each tab in the ID-name map by its title camelCased", () => {
    Object.values(idNameMap.tabs).forEach((tab) => {
      expect(tab.name).toBe(Str.sentenceToCamelCase(tab.title));
    });
  });

  it("place each Table in the ID-name map on a tab that lists it", () => {
    Object.entries(idNameMap.tables).forEach(([tableId, table]) => {
      expect(idNameMap.tabs[String(table.sheetGid)]?.tableIds).toContain(
        tableId,
      );
    });
  });
});

describe("parseForcedTraits", () => {
  it("reads each --force as one table, column and trait", () => {
    expect(
      parseForcedTraits([
        "--send",
        "--force",
        "item/name/header",
        "--force",
        "item/count/columnType",
      ]),
    ).toEqual([
      { table: "item", column: "name", trait: "header" },
      { table: "item", column: "count", trait: "columnType" },
    ]);
  });

  it("forces nothing without the flag", () => {
    expect(parseForcedTraits(["--send"])).toEqual([]);
  });

  it("rejects an address that isn't table/column/trait", () => {
    expect(() => parseForcedTraits(["--force", "item/name"])).toThrow(
      '--force takes <table>/<column>/<trait>, got "item/name".',
    );
  });
});

function enteredDataItem(rowIndex: number, column = "notes"): HeldItem {
  return {
    kind: "enteredData",
    requestIndex: 0,
    verb: "updateCells",
    cell: {
      sheetGid: 7,
      table: "item",
      column,
      row: { rowId: `itm-${rowIndex}` },
    },
    change: { kind: "overwrite", before: `note ${rowIndex}`, after: "=1" },
  };
}

describe("the held file", () => {
  const heldDir = mkdtempSync(join(tmpdir(), "sync-held-"));
  const heldSheetsConfig: SheetsConfig = { ...runSheetsConfig, dir: heldDir };
  const notesItems = Array.from({ length: 23 }, (_, rowIndex) =>
    enteredDataItem(rowIndex),
  );
  const flushes = [[], [...notesItems, enteredDataItem(0, "title")]];
  const forceArgv = ["--force", "item/notes/valueName"];
  const forcedLines = [
    "== Forced: sent over the live value, no --confirm needed ==",
    "item › notes › valueName  sent over the live value by --force",
    "  undone only by the spreadsheet's version history: the send snapshot can't hold it",
  ].join("\n");

  function runner(argv: string[]): SyncRunner {
    return SyncRunner.init(heldSheetsConfig, argv);
  }
  // The gate's held list, its snapshot plans leaving nothing to version history.
  function heldList(token: string, held: HeldItem[][]): HeldList {
    return { token, flushes: held, versionHistoryItems: held.map(() => []) };
  }
  const heldFilePath = runner([]).heldFile.path;

  beforeEach(() => {
    rmSync(join(heldDir, ".probe"), { recursive: true, force: true });
  });

  it("sits in the package's .probe", () => {
    expect(heldFilePath).toBe(join(heldDir, ".probe", "sync-held.json"));
  });

  it("gets every held item and its value from a dry run, with the token", async () => {
    await runner([])._withHeld(
      { ...dryRunResult, outcome: { kind: "ready" } },
      heldList("7f3a9c", flushes),
    );

    expect(JSON.parse(readFileSync(heldFilePath, "utf8"))).toEqual({
      target: "dev",
      token: "7f3a9c",
      flushes,
    });
  });

  it("gives the dry run's report its held list, capped at 20 lines per column, and its path", async () => {
    const result = await runner([])._withHeld(
      { ...dryRunResult, outcome: { kind: "ready" } },
      heldList("7f3a9c", flushes),
    );

    expect(result).toMatchObject({
      held: {
        block: {
          token: "7f3a9c",
          fullListPath: relative(process.cwd(), heldFilePath),
        },
      },
    });
    const text = SyncRunReport.init(result).text;
    expect(text).toContain(
      "== Held: 24 cells: 23 + 1, sent only with --confirm ==",
    );
    expect(text).toContain('  itm-19  "note 19"  → "=1"');
    expect(text).not.toContain("itm-20");
    expect(text).toContain("  … and 3 more in notes, counted, not listed");
    expect(text).toContain(
      `Full list, every cell with its value: ${relative(process.cwd(), heldFilePath)}`,
    );
    expect(text).toContain("Rerun with --send --confirm 7f3a9c to send these.");
  });

  it("is left alone by a dry run that holds nothing", async () => {
    const result = await runner([])._withHeld(
      { ...dryRunResult, outcome: { kind: "ready" } },
      heldList("e3b0c4", [[], []]),
    );

    expect(result).not.toHaveProperty("held");
    expect(existsSync(heldFilePath)).toBe(false);
  });

  it("is left alone by a --force dry run that holds nothing, which prints the forced trait and no held block", async () => {
    const result = await runner(forceArgv)._withHeld(
      { ...dryRunResult, outcome: { kind: "ready" } },
      heldList("e3b0c4", [[], []]),
    );

    const text = SyncRunReport.init(result).text;
    expect(text).toContain(forcedLines);
    expect(text).not.toContain("== Held");
    expect(text).toContain("READY: nothing held");
    expect(text).not.toContain("--confirm e3b0c4");
    expect(existsSync(heldFilePath)).toBe(false);
  });

  it("prints a --force dry run's forced trait beside the held block, outside its count", async () => {
    const result = await runner(forceArgv)._withHeld(
      { ...dryRunResult, outcome: { kind: "ready" } },
      heldList("7f3a9c", flushes),
    );

    const text = SyncRunReport.init(result).text;
    expect(text).toContain(forcedLines);
    expect(text).toContain(
      "== Held: 24 cells: 23 + 1, sent only with --confirm ==",
    );
    expect(text).toContain("READY, 24 HELD");
    expect(JSON.parse(readFileSync(heldFilePath, "utf8"))).toEqual({
      target: "dev",
      token: "7f3a9c",
      flushes,
    });
  });

  it("is left alone by a send, which prints no held list", async () => {
    const result = await runner(["--send"])._withHeld(
      { ...sendResult, outcome: { kind: "sent", summary: "1 sent" } },
      heldList("7f3a9c", flushes),
    );

    expect(result).not.toHaveProperty("held");
    expect(existsSync(heldFilePath)).toBe(false);
  });

  it("is read back by a send with --confirm, and passed with the token", async () => {
    await runner([])._withHeld(
      { ...dryRunResult, outcome: { kind: "ready" } },
      heldList("7f3a9c", flushes),
    );

    expect(
      runner(["--send", "--confirm", "7f3a9c"]).heldFile.confirmedList,
    ).toEqual({ token: "7f3a9c", flushes });
  });

  it("counts as mismatched when missing or unreadable", () => {
    const { heldFile } = runner(["--send", "--confirm", "7f3a9c"]);
    expect(heldFile.confirmedList).toEqual({
      token: "7f3a9c",
      flushes: undefined,
    });
    mkdirSync(dirname(heldFilePath), { recursive: true });
    writeFileSync(heldFilePath, "{ not json");
    expect(heldFile.flushes).toBeUndefined();
    writeFileSync(heldFilePath, JSON.stringify({ flushes: "none" }));
    expect(heldFile.flushes).toBeUndefined();
  });

  it("isn't read by a send without --confirm or by a dry run", () => {
    expect(runner(["--send"]).heldFile.confirmedList).toBeUndefined();
    expect(
      runner(["--confirm", "7f3a9c"]).heldFile.confirmedList,
    ).toBeUndefined();
  });
});

describe("a refused send's report", () => {
  beforeEach(() => {
    vi.spyOn(console, "error").mockImplementation(() => undefined);
  });
  afterEach(() => {
    process.exitCode = undefined;
    vi.restoreAllMocks();
  });

  async function printedReport(result: SyncRunResult): Promise<string> {
    const log = vi.spyOn(console, "log").mockImplementation(() => undefined);
    await exitCodeOf({ run: { result } });
    return log.mock.calls.map(([line]) => String(line)).join("\n");
  }

  it("says only to rerun the dry run when refused before any flush landed", async () => {
    const text = await printedReport({
      ...sendResult,
      outcome: { kind: "stoppedBeforeSending", heldBy: "uncovered" },
    });

    expect(text).toContain(
      "STOPPED: the send was refused before any flush landed",
    );
    expect(text).toContain("Rerun the dry run: npm run sync");
    expect(text).not.toContain("--confirm ");
    expect(text).not.toContain("Wrote ");
  });

  it("reports the landed flush without a count when there was no held file", async () => {
    const text = await printedReport({
      ...sendResult,
      outcome: {
        kind: "stoppedPartway",
        landedFlushes: 1,
        heldBy: "mismatched",
      },
    });

    expect(text).toContain(
      "STOPPED PARTWAY: flush 1 landed; flush 2 was refused",
    );
    expect(text).toContain("Rerun the dry run: npm run sync");
    expect(text).not.toContain("Wrote ");
  });
});
