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
import {
  alterationsFileSource,
  columnIdsByName,
} from "../src/05_Operators/configFileSource.ts";
import { parseIdNameMapFileSource } from "../src/05_Operators/parseIdNameMapFileSource.ts";
import type {
  ConfigRegeneration,
  SyncRun,
} from "../src/05_Operators/SyncCoordinator.ts";
import type { HeldItem } from "../src/nodeHost/FlushClassifier/flushClassifierTypes.ts";
import { SyncRunReport } from "../src/nodeHost/SyncRunReport.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";
import {
  alterationsPath,
  alterationsText,
  appConfigsText,
  ConfigFilesGenerator,
  heldFilePath,
  idNameMapPath,
  loadIntendedState,
  loadPreviousIdNameMap,
  parseConfirmToken,
  parseForcedTraits,
  readHeldFlushes,
} from "./sync.ts";

const devGeneratedDir = fileURLToPath(
  new URL("../dev/generated", import.meta.url),
);
const devIdNameMapPath = idNameMapPath(devGeneratedDir);
const devAlterationsPath = alterationsPath(devGeneratedDir);
const missingFilesDir = mkdtempSync(join(tmpdir(), "sync-"));
const fixtureAlterationsDir = mkdtempSync(join(tmpdir(), "sync-"));

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
  const generator = ConfigFilesGenerator.init(runSheetsConfig, []);
  vi.spyOn(generator, "_sync").mockResolvedValue(run);
  vi.spyOn(generator, "_runTsc").mockReturnValue(isTscPassing);
  await generator.run();
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
    const held = { groups: [], fullListPath: "held.txt", token: "abc" };
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

describe("appConfigsText", () => {
  const text = appConfigsText();

  it("builds appConfigs from the sibling config files", () => {
    expect(text).toContain('import { columnConfigs } from "./columnConfigs";');
    expect(text).toContain(
      'import { retiredIdPrefixes } from "./retiredIdPrefixes";',
    );
    expect(text).toContain('import { tabConfigs } from "./tabConfigs";');
    expect(text).toContain(
      'import { conditionalFormatConfigs } from "./conditionalFormatConfigs";',
    );
    expect(text).toContain('import { tableConfigs } from "./tableConfigs";');
    expect(text).toContain('import { valueConfigs } from "./valueConfigs";');
    expect(text).toContain(
      "export const appConfigs = {\n  tableConfigs,\n  tabConfigs,\n  conditionalFormatConfigs,\n  retiredIdPrefixes,\n  columnConfigs,\n  valueConfigs,\n};",
    );
  });

  it("brings the framework's entry into the program with a type-only import", () => {
    expect(text).toContain(
      'import type {} from "@byronbroughten/sheets-framework";',
    );
  });

  it("augments the package's Register with the appConfigs type", () => {
    expect(text).toMatch(
      /declare module "@byronbroughten\/sheets-framework" \{\s*interface Register \{\s*configs: typeof appConfigs;\s*idNameMap: IdNameMap;\s*\}\s*\}/,
    );
  });

  it("registers the sibling ID-name map's type", () => {
    expect(text).toContain('import type { IdNameMap } from "./idNameMap";');
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

describe("parseConfirmToken", () => {
  it("reads the token after --confirm", () => {
    expect(parseConfirmToken(["--send", "--confirm", "7f3a9c"])).toBe("7f3a9c");
  });

  it("is undefined without the flag", () => {
    expect(parseConfirmToken(["--send"])).toBeUndefined();
  });

  it("rejects --confirm with no token", () => {
    expect(() => parseConfirmToken(["--send", "--confirm"])).toThrow(
      "--confirm takes the token the dry run printed.",
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
  const heldFile = heldFilePath(heldDir);
  const notesItems = Array.from({ length: 23 }, (_, rowIndex) =>
    enteredDataItem(rowIndex),
  );
  const flushes = [[], [...notesItems, enteredDataItem(0, "title")]];

  function generator(argv: string[]): ConfigFilesGenerator {
    return ConfigFilesGenerator.init(heldSheetsConfig, argv);
  }

  beforeEach(() => {
    rmSync(join(heldDir, ".probe"), { recursive: true, force: true });
  });

  it("sits in the package's .probe", () => {
    expect(heldFile).toBe(join(heldDir, ".probe", "sync-held.json"));
  });

  it("gets every held item and its value from a dry run, with the token", async () => {
    await generator([])._withHeld(
      { ...dryRunResult, outcome: { kind: "ready" } },
      { token: "7f3a9c", flushes },
    );

    expect(JSON.parse(readFileSync(heldFile, "utf8"))).toEqual({
      target: "dev",
      token: "7f3a9c",
      flushes,
    });
  });

  it("gives the dry run's report its held list, capped at 20 lines per column, and its path", async () => {
    const result = await generator([])._withHeld(
      { ...dryRunResult, outcome: { kind: "ready" } },
      { token: "7f3a9c", flushes },
    );

    expect(result).toMatchObject({
      held: {
        token: "7f3a9c",
        fullListPath: relative(process.cwd(), heldFile),
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
      `Full list, every cell with its value: ${relative(process.cwd(), heldFile)}`,
    );
    expect(text).toContain("Rerun with --send --confirm 7f3a9c to send these.");
  });

  it("is left alone by a dry run that holds nothing", async () => {
    const result = await generator([])._withHeld(
      { ...dryRunResult, outcome: { kind: "ready" } },
      { token: "e3b0c4", flushes: [[], []] },
    );

    expect(result).not.toHaveProperty("held");
    expect(existsSync(heldFile)).toBe(false);
  });

  it("is left alone by a send, which prints no held list", async () => {
    const result = await generator(["--send"])._withHeld(
      { ...sendResult, outcome: { kind: "sent", summary: "1 sent" } },
      { token: "7f3a9c", flushes },
    );

    expect(result).not.toHaveProperty("held");
    expect(existsSync(heldFile)).toBe(false);
  });

  it("is read back by a send with --confirm, and passed with the token", async () => {
    await generator([])._withHeld(
      { ...dryRunResult, outcome: { kind: "ready" } },
      { token: "7f3a9c", flushes },
    );

    expect(
      generator(["--send", "--confirm", "7f3a9c"])._confirmedList(),
    ).toEqual({ token: "7f3a9c", flushes });
  });

  it("counts as mismatched when missing or unreadable", () => {
    const send = generator(["--send", "--confirm", "7f3a9c"]);
    expect(send._confirmedList()).toEqual({
      token: "7f3a9c",
      flushes: undefined,
    });
    mkdirSync(dirname(heldFile), { recursive: true });
    writeFileSync(heldFile, "{ not json");
    expect(readHeldFlushes(heldFile)).toBeUndefined();
    writeFileSync(heldFile, JSON.stringify({ flushes: "none" }));
    expect(readHeldFlushes(heldFile)).toBeUndefined();
  });

  it("isn't read by a send without --confirm or by a dry run", () => {
    expect(generator(["--send"])._confirmedList()).toBeUndefined();
    expect(generator(["--confirm", "7f3a9c"])._confirmedList()).toBeUndefined();
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

describe("alterationsPath", () => {
  it("puts alterations.ts beside the generated directory", () => {
    expect(alterationsPath(join("app", "src", "generated"))).toBe(
      join("app", "src", "alterations.ts"),
    );
  });
});

describe("loadIntendedState", () => {
  it("reads a missing alterations.ts as an empty one", async () => {
    const generatedDir = join(missingFilesDir, "generated");
    expect(await loadIntendedState(generatedDir)).toEqual({ alterations: {} });
  });

  it("takes the pending alterations alterations.ts exports", async () => {
    const generatedDir = join(fixtureAlterationsDir, "generated");
    writeFileSync(
      alterationsPath(generatedDir),
      "export const pendingAlterations = { tables: { item: { delete: true } } };\n",
    );
    expect(await loadIntendedState(generatedDir)).toEqual({
      alterations: { tables: { item: { delete: true } } },
    });
  });
});

describe("loadPreviousIdNameMap", () => {
  it("reads a missing ID-name map as an empty map", () => {
    expect(loadPreviousIdNameMap(missingFilesDir)).toEqual({
      tables: {},
      tabs: {},
    });
  });

  it("reads the ID-name map back from its source", () => {
    const { tables } = loadPreviousIdNameMap(devGeneratedDir);
    for (const [tableKey, { tableId }] of Object.entries(
      appConfigs.tableConfigs,
    )) {
      expect(tables[tableId]?.name).toBe(tableKey);
    }
  });
});

describe("alterationsText", () => {
  it("is the empty alterations file source", () => {
    expect(alterationsText()).toBe(alterationsFileSource({}));
  });

  it("matches the dev package's own alterations.ts", () => {
    expect(readFileSync(devAlterationsPath, "utf8")).toBe(alterationsText());
  });
});
