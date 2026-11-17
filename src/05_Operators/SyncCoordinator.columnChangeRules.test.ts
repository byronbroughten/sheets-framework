import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SheetIndex } from "../00_Source/RawSource/SheetIndex";
import {
  BraceFormula,
  type RuleLayout,
} from "../01_SpreadsheetSchema/configReaders/BraceFormula";
import { type PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  type FakeSheetsService,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import {
  itemColumns,
  itemGid,
  itemSheet,
  noteColIndex,
  runItem,
} from "./syncCoordinatorTestSupport";

type SheetsRule = GoogleAppsScript.Sheets.Schema.ConditionalFormatRule;
type GridRange = GoogleAppsScript.Sheets.Schema.GridRange;

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("SyncCoordinator.run rules after a column insert or delete", () => {
  const sheet = itemSheet({
    bodyRows: [{ id: "r:itm:one" }, { id: "r:itm:two" }, { id: "r:itm:three" }],
  });
  const itemTable = Val.assert(sheet.tables?.[0], "the Item Table");
  const firstColIndex = Val.assert(itemTable.startColumnIndex, "first col");
  const headerRowIndex = Val.assert(itemTable.startRowIndex, "header row");
  const endRowIndex = Val.assert(itemTable.endRowIndex, "end row");
  const columnNames = Object.keys(itemColumns);
  const countColIndex = columnNames.indexOf("requiredCount");
  const lastColumn = Val.assert(columnNames.at(-1), "last column");
  // Right of both changes, so the fake leaves its anchored reference on the column's old letter.
  const wholeTableFormula = `=\${@${lastColumn}}=""`;
  const countFormula = `=\${@requiredCount}=""`;

  function layoutOf(names: string[]): RuleLayout {
    return {
      headerTopLeft: {
        rowIndex: SheetIndex.row(headerRowIndex),
        colIndex: SheetIndex.col(firstColIndex),
      },
      endRowIndex: SheetIndex.row(endRowIndex),
      columnNamesLeftToRight: names,
    };
  }

  function rangeOver(colOffset: number, width: number): GridRange {
    return {
      sheetId: itemGid,
      startRowIndex: headerRowIndex + 1,
      endRowIndex,
      startColumnIndex: firstColIndex + colOffset,
      endColumnIndex: firstColIndex + colOffset + width,
    };
  }

  function formulaRule(range: GridRange, formula: string): SheetsRule {
    return {
      ranges: [range],
      booleanRule: {
        condition: {
          type: "CUSTOM_FORMULA",
          values: [{ userEnteredValue: formula }],
        },
        format: { backgroundColor: { red: 1, green: 0, blue: 0 } },
      },
    };
  }

  // Over every column, so the fake replays it across the change point.
  function wholeTableRule(names: string[]): SheetsRule {
    return formulaRule(
      rangeOver(0, names.length),
      BraceFormula.compileForRule({
        tableKey: "item",
        formula: wholeTableFormula,
        layout: layoutOf(names),
      }),
    );
  }

  // Left of the count column, reading it: the fake throws on its delete unless the rule went first.
  const countReadingRule = formulaRule(
    rangeOver(noteColIndex, 1),
    BraceFormula.compileForRule({
      tableKey: "item",
      formula: countFormula,
      layout: layoutOf(columnNames),
    }),
  );
  const countOnlyRule = formulaRule(rangeOver(countColIndex, 1), "=TRUE");
  // From the header row, so it has no brace form.
  const handDrawnRule = formulaRule(
    { ...rangeOver(noteColIndex, 1), startRowIndex: headerRowIndex },
    "=TRUE",
  );
  const countReadingHandle = `optionalNote | ${countFormula} | background #ff0000`;
  const countOnlyHandle = "requiredCount | =TRUE | background #ff0000";

  function seedWith(rules: SheetsRule[]): FakeSheetsService {
    return stubSheetsService({
      sheets: [{ ...sheet, conditionalFormats: rules }],
    });
  }

  function rulesLeft(service: FakeSheetsService): SheetsRule[] {
    return service.grid.sheet(itemGid).conditionalFormats ?? [];
  }

  function runWith(isSending: boolean, alterations: PendingAlterations) {
    return runItem({ isSending, alterations });
  }

  const deleteCount: PendingAlterations = {
    tables: { item: { columns: { requiredCount: { deleted: true } } } },
  };
  const insertRef: PendingAlterations = {
    tables: {
      item: { newColumns: [{ header: "Ref", before: "requiredCount" }] },
    },
  };

  it("after a delete, rewrites a whole-Table rule, removes one reading the column and leaves one wholly over it to Sheets", () => {
    const service = seedWith([
      wholeTableRule(columnNames),
      countReadingRule,
      countOnlyRule,
      handDrawnRule,
    ]);

    const { result } = runWith(true, deleteCount);

    const namesAfter = columnNames.filter((name) => name !== "requiredCount");
    expect(rulesLeft(service)).toEqual([
      wholeTableRule(namesAfter),
      handDrawnRule,
    ]);
    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      stops: [],
      didntLand: [],
    });
    expect(result.plan).toEqual([
      {
        subject: { table: "item", column: "requiredCount", path: "deleted" },
        what: "delete the column",
        effects: [
          {
            kind: "rulesRemoved",
            handles: [countReadingHandle, countOnlyHandle],
          },
          { kind: "rulesRewritten", table: "item", count: 1, after: "delete" },
        ],
      },
    ]);
  });

  it("after an insert, rewrites a whole-Table rule over the new column too", () => {
    const service = seedWith([wholeTableRule(columnNames), handDrawnRule]);

    const { result } = runWith(true, insertRef);

    const namesAfter = [...columnNames];
    namesAfter.splice(countColIndex, 0, "ref");
    expect(rulesLeft(service)).toEqual([
      wholeTableRule(namesAfter),
      handDrawnRule,
    ]);
    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      stops: [],
      didntLand: [],
    });
    expect(
      result.plan.flatMap(({ effects }) =>
        effects.filter(({ kind }) => kind === "rulesRewritten"),
      ),
    ).toEqual([
      { kind: "rulesRewritten", table: "item", count: 1, after: "insert" },
    ]);
  });

  it("lists a hand-drawn rule on a dry run, by its handle", () => {
    seedWith([handDrawnRule]);

    const { result } = runWith(false, insertRef);

    expect(result.checks).toEqual([
      expect.objectContaining({
        kind: "handDrawnRule",
        table: "item",
        formula: "=TRUE",
        handle: expect.stringContaining("unmodelable"),
        cause:
          "couldn't be rewritten, and the insert may leave it pointing at the wrong column",
      }),
    ]);
  });

  it("sends no rule after a header rename", () => {
    const rules = [wholeTableRule(columnNames), handDrawnRule];
    const service = seedWith(rules);

    const { result } = runWith(true, {
      tables: { item: { columns: { optionalNote: { header: "Remark" } } } },
    });

    expect(rulesLeft(service)).toEqual(rules);
    expect(result.checks).toEqual([]);
    expect(result.plan.flatMap(({ effects }) => effects)).toEqual([]);
  });
});
