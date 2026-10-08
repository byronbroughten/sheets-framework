import { describe, expect, it } from "vitest";

import { getColumnTraitByName } from "../01_SpreadsheetSchema/columnConfigsTypes";
import { getTableTraitByName } from "../01_SpreadsheetSchema/tableConfigsTypes";
import {
  layoutBodyRows,
  layoutGid,
  layoutSheet,
  type LayoutTableName,
  layoutTableNames,
} from "../02_SpreadsheetRaw/spreadsheetRawTestSupport";
import { expectedSheetLayout } from "../testSupport/expectedSheetLayout";
import {
  buildGridRows,
  type FakeSheetsService,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { Val } from "../utils/Val";
import { SpreadsheetNamed } from "./SpreadsheetNamed";

const topDataRowIndex = expectedSheetLayout.tableHeaderRowIndex + 1;
const runItemGid = getTableTraitByName("runItem", "sheetGid");
const idColumnId = getColumnTraitByName("runItem", "id", "columnId");
const selectColumnId = getColumnTraitByName("runItem", "selected", "columnId");
const pink = { red: 244 / 255, green: 204 / 255, blue: 204 / 255 };
const grey = { red: 0.6, green: 0.6, blue: 0.6 };
const green = { red: 0.7, green: 0.9, blue: 0.7 };

const sheetRange = {
  sheetId: runItemGid,
  startRowIndex: topDataRowIndex,
  endRowIndex: 6,
  startColumnIndex: 0,
  endColumnIndex: 2,
};
const idColumnRange = {
  ...sheetRange,
  startColumnIndex: 0,
  endColumnIndex: 1,
};
const selectColumnRange = {
  ...sheetRange,
  startColumnIndex: 1,
  endColumnIndex: 2,
};
const topIdCellRange = {
  ...idColumnRange,
  endRowIndex: topDataRowIndex + 1,
};

function googleBooleanRule(
  range: GoogleAppsScript.Sheets.Schema.GridRange,
  type: string,
  userEnteredValue: string,
  backgroundColor: { red: number; green: number; blue: number },
): GoogleAppsScript.Sheets.Schema.ConditionalFormatRule {
  return {
    ranges: [range],
    booleanRule: {
      condition: { type, values: [{ userEnteredValue }] },
      format: { backgroundColor },
    },
  };
}

function stubRunItemWithRules(
  conditionalFormats: GoogleAppsScript.Sheets.Schema.ConditionalFormatRule[],
) {
  return stubSheetsService({
    sheets: [
      {
        sheetId: runItemGid,
        title: "Run item",
        rows: buildGridRows({
          0: [idColumnId, selectColumnId],
          3: ["ID", "Selected"],
          4: ["r:rit:row4", true],
          5: [null, null],
        }),
        table: { endRowIndex: 6, endColumnIndex: 2 },
        conditionalFormats,
      },
    ],
  });
}

function runItemRules(grid: FakeSheetsService["grid"]) {
  return grid.sheet(runItemGid).conditionalFormats;
}

function fetchedRunItem() {
  const ss = SpreadsheetNamed.init();
  const sheet = ss.table("runItem");
  sheet.sheet.prepFetchConditionalFormatRules();
  ss.fetchAllPrepped();
  return { ss, sheet };
}

describe("TableNamed conditional format rules", () => {
  it("reads a sheet with no rules as an empty list", () => {
    stubRunItemWithRules([]);
    const { sheet } = fetchedRunItem();

    expect(sheet.sheet.conditionalFormatRules()).toEqual([]);
  });

  it("prepends a column rule so it takes precedence over rules already on the sheet", () => {
    const { grid } = stubRunItemWithRules([
      googleBooleanRule(sheetRange, "NUMBER_EQ", "TRUE", grey),
    ]);
    const { ss, sheet } = fetchedRunItem();

    sheet.column("id").addConditionalFormatRule({
      condition: { type: "NUMBER_EQ", value: true },
      format: { backgroundColor: pink },
    });
    ss.batchUpdateGSheets();

    expect(runItemRules(grid)).toEqual([
      googleBooleanRule(idColumnRange, "NUMBER_EQ", "TRUE", pink),
      googleBooleanRule(sheetRange, "NUMBER_EQ", "TRUE", grey),
    ]);

    sheet.sheet.prepFetchConditionalFormatRules();
    ss.fetchAllPrepped({ skipFetchingProperties: true });
    expect(sheet.sheet.conditionalFormatRules()[0]).toMatchObject({
      kind: "boolean",
      condition: { type: "NUMBER_EQ", value: true },
    });
    expect(sheet.sheet.conditionalFormatRules()[1]).toMatchObject({
      kind: "boolean",
      format: { backgroundColor: grey },
    });
  });

  it("removes every rule whose range exactly matches a column and leaves a sheet-wide rule alone", () => {
    const { grid } = stubRunItemWithRules([
      googleBooleanRule(sheetRange, "NUMBER_EQ", "TRUE", grey),
      googleBooleanRule(idColumnRange, "NUMBER_EQ", "TRUE", pink),
      googleBooleanRule(selectColumnRange, "NUMBER_EQ", "TRUE", green),
      googleBooleanRule(idColumnRange, "NUMBER_NOT_EQ", "TRUE", green),
    ]);
    const { ss, sheet } = fetchedRunItem();

    sheet.column("id").removeConditionalFormatRules();
    ss.batchUpdateGSheets();

    expect(runItemRules(grid)).toEqual([
      googleBooleanRule(sheetRange, "NUMBER_EQ", "TRUE", grey),
      googleBooleanRule(selectColumnRange, "NUMBER_EQ", "TRUE", green),
    ]);

    sheet.sheet.prepFetchConditionalFormatRules();
    ss.fetchAllPrepped({ skipFetchingProperties: true });
    const rules = sheet.sheet.conditionalFormatRules();
    expect(rules).toHaveLength(2);
    expect(rules[0]).toMatchObject({
      ranges: [sheetRange],
      format: { backgroundColor: grey },
    });
    expect(rules[1]).toMatchObject({
      ranges: [selectColumnRange],
    });
  });

  it("removes a single rule by exact content so a re-stamp can replace it", () => {
    const { grid } = stubRunItemWithRules([
      googleBooleanRule(idColumnRange, "NUMBER_EQ", "TRUE", pink),
      googleBooleanRule(idColumnRange, "NUMBER_NOT_EQ", "TRUE", green),
    ]);
    const { ss, sheet } = fetchedRunItem();
    const toRemove = Val.assert(
      sheet.sheet.conditionalFormatRules()[0],
      "ID column rule to remove",
    );

    sheet.column("id").removeConditionalFormatRule(toRemove);
    ss.batchUpdateGSheets();

    expect(runItemRules(grid)).toEqual([
      googleBooleanRule(idColumnRange, "NUMBER_NOT_EQ", "TRUE", green),
    ]);
  });

  it("sends no batch update when adding a rule identical to one already present", () => {
    const { batchUpdateCount } = stubRunItemWithRules([
      googleBooleanRule(idColumnRange, "NUMBER_EQ", "TRUE", pink),
    ]);
    const { ss, sheet } = fetchedRunItem();

    sheet.column("id").addConditionalFormatRule({
      condition: { type: "NUMBER_EQ", value: true },
      format: { backgroundColor: pink },
    });
    ss.batchUpdateGSheets();

    expect(batchUpdateCount()).toBe(0);
  });

  it("removes several column rules at once and keeps the rule between them", () => {
    const { grid } = stubRunItemWithRules([
      googleBooleanRule(idColumnRange, "NUMBER_EQ", "1", pink),
      googleBooleanRule(sheetRange, "NUMBER_EQ", "TRUE", grey),
      googleBooleanRule(idColumnRange, "NUMBER_EQ", "2", green),
      googleBooleanRule(idColumnRange, "NUMBER_EQ", "3", grey),
    ]);
    const { ss, sheet } = fetchedRunItem();

    sheet.column("id").removeConditionalFormatRules();
    ss.batchUpdateGSheets();

    expect(runItemRules(grid)).toEqual([
      googleBooleanRule(sheetRange, "NUMBER_EQ", "TRUE", grey),
    ]);
  });

  it("refuses a second rule-mutating flush against the same sheet until the rules are re-fetched", () => {
    const { grid } = stubRunItemWithRules([]);
    const { ss, sheet } = fetchedRunItem();

    sheet.column("id").addConditionalFormatRule({
      condition: { type: "NUMBER_EQ", value: true },
      format: { backgroundColor: pink },
    });
    ss.batchUpdateGSheets();

    expect(() =>
      sheet.column("id").addConditionalFormatRule({
        condition: { type: "NUMBER_NOT_EQ", value: true },
        format: { backgroundColor: green },
      }),
    ).toThrowError(/Conditional format indexes are stale/);

    sheet.sheet.prepFetchConditionalFormatRules();
    ss.fetchAllPrepped({ skipFetchingProperties: true });
    sheet.column("id").addConditionalFormatRule({
      condition: { type: "NUMBER_NOT_EQ", value: true },
      format: { backgroundColor: green },
    });
    ss.batchUpdateGSheets();
    expect(runItemRules(grid)).toEqual([
      googleBooleanRule(idColumnRange, "NUMBER_NOT_EQ", "TRUE", green),
      googleBooleanRule(idColumnRange, "NUMBER_EQ", "TRUE", pink),
    ]);
  });

  it("builds an anchored A1 reference from a column name at the data-start row", () => {
    stubRunItemWithRules([]);
    const { sheet } = fetchedRunItem();

    expect(sheet.column("id").anchoredA1()).toBe("$A5");
    expect(sheet.column("id").anchoredA1("selected")).toBe("$B5");
    expect(sheet.column("id").cell(1).anchoredA1("id")).toBe("$A6");
  });

  it("adds a sheet-wide rule over the live data range and a cell rule over one cell", () => {
    const { grid } = stubRunItemWithRules([]);
    const { ss, sheet } = fetchedRunItem();

    sheet.addConditionalFormatRule({
      condition: { type: "NUMBER_EQ", value: true },
      format: { backgroundColor: grey, foregroundColor: grey },
    });
    sheet
      .column("id")
      .cell(0)
      .addConditionalFormatRule({
        condition: {
          type: "CUSTOM_FORMULA",
          formula: `=${sheet.column("id").cell(0).anchoredA1()}=FALSE`,
        },
        format: { backgroundColor: pink },
      });
    ss.batchUpdateGSheets();

    expect(runItemRules(grid).map((rule) => rule.ranges)).toEqual([
      [topIdCellRange],
      [sheetRange],
    ]);
  });

  it("replaces a column rule in one batch update, so a malformed add cannot leave the column bare", () => {
    const { batchUpdateCount, grid } = stubRunItemWithRules([
      googleBooleanRule(idColumnRange, "CUSTOM_FORMULA", "=$B5=TRUE", green),
    ]);
    const { ss, sheet } = fetchedRunItem();
    const existing = Val.assert(
      sheet.sheet.conditionalFormatRules()[0],
      "existing ID column rule",
    );
    const idPrefix = sheet.column("id");

    idPrefix.removeConditionalFormatRule(existing);
    idPrefix.addConditionalFormatRule({
      condition: {
        type: "CUSTOM_FORMULA",
        formula: `=${idPrefix.anchoredA1("selected")}=FALSE`,
      },
      format: { backgroundColor: pink },
    });
    ss.batchUpdateGSheets();

    expect(batchUpdateCount()).toBe(1);
    expect(runItemRules(grid)).toEqual([
      googleBooleanRule(idColumnRange, "CUSTOM_FORMULA", "=$B5=FALSE", pink),
    ]);
  });
});

function googleProtection(
  range: GoogleAppsScript.Sheets.Schema.GridRange,
  extras: Partial<GoogleAppsScript.Sheets.Schema.ProtectedRange> = {},
): GoogleAppsScript.Sheets.Schema.ProtectedRange {
  return {
    protectedRangeId: extras.protectedRangeId ?? 1,
    range,
    warningOnly: extras.warningOnly ?? true,
    ...extras,
  };
}

function stubRunItemWithProtections(
  protectedRanges: GoogleAppsScript.Sheets.Schema.ProtectedRange[],
) {
  return stubSheetsService({
    sheets: [
      {
        sheetId: runItemGid,
        title: "Run item",
        rows: buildGridRows({
          0: [idColumnId, selectColumnId],
          3: ["ID", "Selected"],
          4: ["r:rit:row4", true],
          5: [null, null],
        }),
        table: { endRowIndex: 6, endColumnIndex: 2 },
        protectedRanges,
      },
    ],
  });
}

function fetchedRunItemProtections(
  protectedRanges: GoogleAppsScript.Sheets.Schema.ProtectedRange[] = [],
) {
  const service = stubRunItemWithProtections(protectedRanges);
  const ss = SpreadsheetNamed.init();
  const sheet = ss.table("runItem");
  sheet.sheet.prepFetchEditProtections();
  ss.fetchAllPrepped();
  const protections = () => service.grid.sheet(runItemGid).protectedRanges;
  return { ss, sheet, protections, ...service };
}

const wholeSheetRange = { sheetId: runItemGid };
const columnIdRowRange = {
  sheetId: runItemGid,
  startRowIndex: 0,
  endRowIndex: 1,
};
const idHeaderCellRange = {
  sheetId: runItemGid,
  startRowIndex: 3,
  endRowIndex: 4,
  startColumnIndex: 0,
  endColumnIndex: 1,
};
const idColumnIdCellRange = {
  sheetId: runItemGid,
  startRowIndex: 0,
  endRowIndex: 1,
  startColumnIndex: 0,
  endColumnIndex: 1,
};
const idGroupHeadingCellRange = {
  sheetId: runItemGid,
  startRowIndex: 1,
  endRowIndex: 2,
  startColumnIndex: 0,
  endColumnIndex: 1,
};
const idWholeColumnRange = {
  sheetId: runItemGid,
  startRowIndex: 0,
  startColumnIndex: 0,
  endColumnIndex: 1,
};
const idWholeColumnGoogleRange = {
  sheetId: runItemGid,
  startColumnIndex: 0,
  endColumnIndex: 1,
};

describe("TableNamed edit warnings and edit locks", () => {
  it("sends no batch update when adding a warning identical to one already present", () => {
    const { batchUpdateCount, ss, sheet } = fetchedRunItemProtections([
      googleProtection(idColumnRange, {
        protectedRangeId: 4,
        description: "id warning",
        warningOnly: true,
      }),
    ]);

    sheet.column("id").addEditWarning({ description: "id warning" });
    ss.batchUpdateGSheets();

    expect(batchUpdateCount()).toBe(0);
  });

  it("sends no batch update when a present lock has every declared editor plus ones Google added", () => {
    const googleAdded = ["service@example.com", "owner@example.com"];
    const fetchedLock = (users: string[]) =>
      fetchedRunItemProtections([
        googleProtection(idColumnRange, {
          protectedRangeId: 6,
          description: "id lock",
          warningOnly: false,
          editors: { users, groups: ["editors@example.com"] },
        }),
      ]);

    const unnamed = fetchedLock(googleAdded);
    unnamed.sheet.column("id").addEditLock({ description: "id lock" });
    unnamed.ss.batchUpdateGSheets();
    expect(unnamed.batchUpdateCount()).toBe(0);

    const named = fetchedLock([...googleAdded, "editor@example.com"]);
    named.sheet.column("id").addEditLock({
      description: "id lock",
      users: ["editor@example.com"],
      groups: ["editors@example.com"],
    });
    named.ss.batchUpdateGSheets();
    expect(named.batchUpdateCount()).toBe(0);
  });

  it("adds a lock over a head cell", () => {
    const { protections, ss, sheet } = fetchedRunItemProtections();

    sheet.column("id").headCell("header").addEditLock({
      description: "id header lock",
    });
    ss.batchUpdateGSheets();

    expect(protections()).toHaveLength(1);
    expect(protections()[0]).toMatchObject({
      range: idHeaderCellRange,
      description: "id header lock",
    });
    expect(protections()[0]?.warningOnly).not.toBe(true);
  });

  it("adds a lock when a present lock lacks a declared editor", () => {
    const { protections, ss, sheet } = fetchedRunItemProtections([
      googleProtection(idColumnRange, {
        protectedRangeId: 6,
        description: "id lock",
        warningOnly: false,
        editors: { users: ["service@example.com", "owner@example.com"] },
      }),
    ]);

    sheet.column("id").addEditLock({
      description: "id lock",
      users: ["editor@example.com"],
    });
    ss.batchUpdateGSheets();

    expect(protections()).toHaveLength(2);
    expect(protections()[1]).toMatchObject({
      range: idColumnRange,
      editors: { users: ["editor@example.com"] },
    });
  });

  it("removes a hand-set protection by exact range, content, description and id", () => {
    const handSet = googleProtection(idColumnRange, {
      protectedRangeId: 11,
      description: "hand-set",
      warningOnly: true,
    });
    const other = googleProtection(selectColumnRange, {
      protectedRangeId: 12,
      description: "other",
      warningOnly: true,
    });

    const byRange = fetchedRunItemProtections([handSet, other]);
    byRange.sheet.column("id").removeEditProtections();
    byRange.ss.batchUpdateGSheets();
    expect(byRange.protections()).toEqual([other]);

    const byContent = fetchedRunItemProtections([handSet, other]);
    const named = Val.assert(
      byContent.sheet.sheet.editProtections()[0],
      "hand-set protection",
    );
    byContent.sheet.sheet.removeEditProtection(named);
    byContent.ss.batchUpdateGSheets();
    expect(byContent.protections()).toEqual([other]);

    const byDescription = fetchedRunItemProtections([handSet, other]);
    byDescription.sheet.sheet.removeEditProtectionByDescription("hand-set");
    byDescription.ss.batchUpdateGSheets();
    expect(byDescription.protections()).toEqual([other]);

    const byId = fetchedRunItemProtections([handSet, other]);
    byId.sheet.sheet.removeEditProtectionById(11);
    byId.ss.batchUpdateGSheets();
    expect(byId.protections()).toEqual([other]);
  });

  it("adds a whole-sheet warning with unprotected ranges", () => {
    const { protections, ss, sheet } = fetchedRunItemProtections();
    const unprotected = {
      sheetId: runItemGid,
      startRowIndex: 2,
      endRowIndex: 3,
      startColumnIndex: 1,
      endColumnIndex: 2,
    };

    sheet.sheet.addEditWarningWholeSheet({
      description: "sheet warning",
      unprotectedRanges: [unprotected],
    });
    ss.batchUpdateGSheets();

    expect(protections()).toEqual([
      {
        protectedRangeId: expect.any(Number),
        range: wholeSheetRange,
        description: "sheet warning",
        warningOnly: true,
        unprotectedRanges: [unprotected],
      },
    ]);
  });

  it("refuses a coordinate-bearing protection write while row indexes are stale", () => {
    const { ss, sheet } = fetchedRunItemProtections();

    sheet.row(1).delete();
    ss.batchUpdateGSheets();

    expect(() => sheet.column("id").addEditWarning()).toThrowError(
      /Row indexes are stale/,
    );
    expect(() =>
      sheet.sheet.addEditWarningWholeSheet({
        unprotectedRanges: [topIdCellRange],
      }),
    ).toThrowError(/Row indexes are stale/);
    expect(() => sheet.sheet.addEditLockWholeSheet()).not.toThrow();
  });

  it("refuses a read or mutation after a protection flush until protections are re-fetched", () => {
    const { protections, ss, sheet } = fetchedRunItemProtections();

    sheet.column("id").addEditWarning({ description: "id warning" });
    ss.batchUpdateGSheets();

    expect(() => sheet.sheet.editProtections()).toThrowError(
      /Edit protections are stale/,
    );
    expect(() =>
      sheet.column("id").addEditLock({ description: "id lock" }),
    ).toThrowError(/Edit protections are stale/);

    sheet.sheet.prepFetchEditProtections();
    ss.fetchAllPrepped({ skipFetchingProperties: true });
    expect(sheet.sheet.editProtections()[0]).toMatchObject({
      kind: "warning",
      description: "id warning",
    });
    sheet.column("id").addEditLock({ description: "id lock" });
    ss.batchUpdateGSheets();
    expect(protections().map((protection) => protection.description)).toEqual([
      "id warning",
      "id lock",
    ]);
  });

  it("adds a lock with named editors", () => {
    const { protections, ss, sheet } = fetchedRunItemProtections();

    sheet.column("id").addEditLock({
      description: "id lock",
      users: ["editor@example.com"],
      groups: ["editors@example.com"],
    });
    ss.batchUpdateGSheets();

    expect(protections()).toEqual([
      {
        protectedRangeId: expect.any(Number),
        range: idColumnRange,
        description: "id lock",
        editors: {
          users: ["editor@example.com"],
          groups: ["editors@example.com"],
        },
      },
    ]);
  });

  it("adds warnings over a head row, header cells, a column-group heading cell and a single cell", () => {
    const { protections, ss, sheet } = fetchedRunItemProtections();

    sheet.headRow("columnId").addEditWarning({
      description: "column id row",
    });
    sheet.column("id").headCell("header").addEditWarning({
      description: "id header",
    });
    sheet.column("id").headCell("columnId").addEditWarning({
      description: "id column id",
    });
    sheet.column("id").headCell("groupHeading1").addEditWarning({
      description: "id group heading",
    });
    sheet.column("id").cell(0).addEditWarning({
      description: "id cell",
    });
    ss.batchUpdateGSheets();

    expect(protections().map((protection) => protection.range)).toEqual([
      columnIdRowRange,
      idHeaderCellRange,
      idColumnIdCellRange,
      idGroupHeadingCellRange,
      topIdCellRange,
    ]);
  });

  it("adds an open-ended column warning from a start row with no end row", () => {
    const { protections, ss, sheet } = fetchedRunItemProtections();

    sheet.column("id").addEditWarningFromRow(0, {
      description: "open-ended id",
    });
    ss.batchUpdateGSheets();

    expect(protections()).toEqual([
      {
        protectedRangeId: expect.any(Number),
        range: {
          sheetId: runItemGid,
          startRowIndex: topDataRowIndex,
          startColumnIndex: 0,
          endColumnIndex: 1,
        },
        description: "open-ended id",
        warningOnly: true,
      },
    ]);
  });

  it("adds a whole-column warning as a start-row-0 column range with no end row", () => {
    const { protections, ss, sheet } = fetchedRunItemProtections();

    sheet.column("id").addEditWarningWholeColumn({
      description: "id column",
    });
    ss.batchUpdateGSheets();

    expect(protections()).toEqual([
      {
        protectedRangeId: expect.any(Number),
        range: idWholeColumnRange,
        description: "id column",
        warningOnly: true,
      },
    ]);
  });

  it("adds a whole-column lock as a start-row-0 column range with no end row", () => {
    const { protections, ss, sheet } = fetchedRunItemProtections();

    sheet.column("id").addEditLockWholeColumn({
      description: "id column lock",
      users: ["editor@example.com"],
    });
    ss.batchUpdateGSheets();

    expect(protections()).toEqual([
      {
        protectedRangeId: expect.any(Number),
        range: idWholeColumnRange,
        description: "id column lock",
        editors: { users: ["editor@example.com"] },
      },
    ]);
  });

  it("sends no batch update when adding a whole-column warning identical to one already present", () => {
    const { batchUpdateCount, ss, sheet } = fetchedRunItemProtections([
      googleProtection(idWholeColumnGoogleRange, {
        protectedRangeId: 20,
        description: "id column",
        warningOnly: true,
      }),
    ]);

    sheet.column("id").addEditWarningWholeColumn({ description: "id column" });
    ss.batchUpdateGSheets();

    expect(batchUpdateCount()).toBe(0);
  });

  it("removes a whole-column protection by exact range, content, description and id, and leaves a whole-sheet protection alone", () => {
    const wholeColumn = googleProtection(idWholeColumnGoogleRange, {
      protectedRangeId: 21,
      description: "id column",
      warningOnly: true,
    });
    const wholeSheet = googleProtection(wholeSheetRange, {
      protectedRangeId: 22,
      description: "sheet",
      warningOnly: true,
    });

    const byRange = fetchedRunItemProtections([wholeColumn, wholeSheet]);
    byRange.sheet.column("id").removeEditProtectionsWholeColumn();
    byRange.ss.batchUpdateGSheets();
    expect(byRange.protections()).toEqual([wholeSheet]);

    const byContent = fetchedRunItemProtections([wholeColumn, wholeSheet]);
    const named = Val.assert(
      byContent.sheet.sheet.editProtections()[0],
      "whole-column protection",
    );
    byContent.sheet.sheet.removeEditProtection(named);
    byContent.ss.batchUpdateGSheets();
    expect(byContent.protections()).toEqual([wholeSheet]);

    const byDescription = fetchedRunItemProtections([wholeColumn, wholeSheet]);
    byDescription.sheet.sheet.removeEditProtectionByDescription("id column");
    byDescription.ss.batchUpdateGSheets();
    expect(byDescription.protections()).toEqual([wholeSheet]);

    const byId = fetchedRunItemProtections([wholeColumn, wholeSheet]);
    byId.sheet.sheet.removeEditProtectionById(21);
    byId.ss.batchUpdateGSheets();
    expect(byId.protections()).toEqual([wholeSheet]);
  });

  it("removes a whole-sheet protection by exact range without matching a whole-column protection", () => {
    const wholeColumn = googleProtection(idWholeColumnGoogleRange, {
      protectedRangeId: 21,
      description: "id column",
      warningOnly: true,
    });
    const { protections, ss, sheet } = fetchedRunItemProtections([
      wholeColumn,
      googleProtection(wholeSheetRange, {
        protectedRangeId: 22,
        description: "sheet",
        warningOnly: true,
      }),
    ]);

    sheet.identified.raw.sheet.removeEditProtectionsAt(wholeSheetRange);
    ss.batchUpdateGSheets();

    expect(protections()).toEqual([wholeColumn]);
  });

  it("allows a whole-column protection write while row indexes are stale", () => {
    const { ss, sheet } = fetchedRunItemProtections();

    sheet.row(1).delete();
    ss.batchUpdateGSheets();

    expect(() =>
      sheet
        .column("id")
        .addEditWarningWholeColumn({ description: "id column" }),
    ).not.toThrow();
    expect(() =>
      sheet.column("id").addEditLockWholeColumn({
        description: "id column lock",
      }),
    ).not.toThrow();
  });
});

const itemGid = getTableTraitByName("item", "sheetGid");

function fetchedItemNames(
  dataRows: readonly (readonly [string | null, string | null])[],
) {
  const service = stubSheetsService({
    sheets: [
      {
        sheetId: itemGid,
        title: "Item",
        rows: buildGridRows({
          0: [
            getColumnTraitByName("item", "id", "columnId"),
            getColumnTraitByName("item", "name", "columnId"),
          ],
          3: ["ID", "Name"],
          ...Object.fromEntries(
            dataRows.map((row, index) => [topDataRowIndex + index, [...row]]),
          ),
        }),
        table: {
          endRowIndex: topDataRowIndex + dataRows.length,
          endColumnIndex: 2,
        },
      },
    ],
  });
  const ss = SpreadsheetNamed.init();
  const sheet = ss.table("item").prepFetchRowIdAndName();
  ss.fetchAllPrepped();
  return { ss, sheet, ...service };
}

describe("TableNamed.rowIdByName", () => {
  it("finds the one row with the name, with its id and row index", () => {
    const { sheet } = fetchedItemNames([
      ["r:itm:aaaaaaa", "Widget A"],
      ["r:itm:bbbbbbb", "Widget B"],
    ]);

    expect(sheet.rowIdByName("Widget B")).toEqual({
      found: "one",
      rowId: "r:itm:bbbbbbb",
      rowIndex: 1,
    });
  });

  it("reports none when no row has the name", () => {
    const { sheet } = fetchedItemNames([["r:itm:aaaaaaa", "Widget A"]]);

    expect(sheet.rowIdByName("Widget B")).toEqual({ found: "none" });
  });

  it("reports many, with the count, when several rows share the name", () => {
    const { sheet } = fetchedItemNames([
      ["r:itm:aaaaaaa", "Widget B"],
      ["r:itm:bbbbbbb", "Widget A"],
      ["r:itm:ccccccc", "Widget B"],
    ]);

    expect(sheet.rowIdByName("Widget B")).toEqual({
      found: "many",
      rowCount: 2,
    });
  });

  it("fills a blank id on the row it found", () => {
    const { ss, sheet, grid } = fetchedItemNames([[null, "Widget A"]]);

    const match = sheet.rowIdByName("Widget A");
    ss.batchUpdateGSheets();

    expect(match).toMatchObject({ found: "one", rowIndex: 0 });
    const rowId = match.found === "one" ? match.rowId : "";
    expect(rowId).toMatch(/^r:itm:/);
    expect(grid.sheet(itemGid).cell(topDataRowIndex, 0)).toBe(rowId);
  });

  it("refuses a blank name, which would match every unnamed row", () => {
    const { sheet } = fetchedItemNames([["r:itm:aaaaaaa", ""]]);

    expect(() => sheet.rowIdByName("")).toThrow(
      'Cannot look up a blank name in the name column of "item".',
    );
  });

  it("is typed only to sheets with both an id and a name column", () => {
    function neverCalled(ss: SpreadsheetNamed): void {
      ss.table("item").rowIdByName("Widget B");
      // @ts-expect-error valueTypes has an id column but no name column.
      ss.table("valueTypes").rowIdByName("Widget B");
      // @ts-expect-error log has neither.
      ss.table("log").rowIdByName("Widget B");
      // @ts-expect-error spreadsheetConfig has neither.
      ss.table("spreadsheetConfig").prepFetchRowIdAndName();
    }
    expect(neverCalled).toBeTypeOf("function");
  });
});

describe("Tables that share a sheet", () => {
  function fetchedLayout(): {
    ss: SpreadsheetNamed;
    grid: FakeSheetsService["grid"];
  } {
    const { grid } = stubSheetsService({ sheets: [layoutSheet()] });
    const ss = SpreadsheetNamed.init();
    layoutTableNames.forEach((tableName) =>
      ss.table(tableName).prepFetchColumnsFull("entry", "amount"),
    );
    ss.fetchAllPrepped();
    return { ss, grid };
  }

  function placedRange(
    grid: FakeSheetsService["grid"],
    tableName: LayoutTableName,
  ): number[] {
    const table = Val.assert(
      grid
        .sheet(layoutGid)
        .tables.find(
          ({ tableId }) =>
            tableId === getTableTraitByName(tableName, "tableId"),
        ),
      `${tableName} on the grid`,
    );
    const range = Val.assert(table.range, `${tableName}'s range`);
    return [
      range.startRowIndex ?? 0,
      range.endRowIndex ?? 0,
      range.startColumnIndex ?? 0,
      range.endColumnIndex ?? 0,
    ];
  }

  it("reads each Table's rows by its own name", () => {
    const { ss } = fetchedLayout();

    expect(
      layoutTableNames.map((tableName) =>
        ss
          .table(tableName)
          .rows.map((row) => [row.value("entry"), row.value("amount")]),
      ),
    ).toEqual(layoutTableNames.map((tableName) => layoutBodyRows[tableName]));
  });

  it("writes to a Table beside another in its own cells", () => {
    const { ss, grid } = fetchedLayout();
    ss.table("layoutRight").row(0).updateValue("amount", 11);
    ss.batchUpdateGSheets();

    expect(grid.sheet(layoutGid).cell(4, 4)).toBe(11);
    expect(grid.sheet(layoutGid).cell(4, 1)).toBe(1);
  });

  it("pushes the Table below down intact when the one above it grows", () => {
    const { ss, grid } = fetchedLayout();
    ss.table("layoutLeft").appendRowWithVals({ entry: "Left four", amount: 4 });
    ss.batchUpdateGSheets();

    expect(placedRange(grid, "layoutLeft")).toEqual([3, 8, 0, 2]);
    expect(placedRange(grid, "layoutBelow")).toEqual([12, 15, 0, 2]);
    expect(placedRange(grid, "layoutRight")).toEqual([3, 6, 3, 5]);
    expect(
      grid
        .sheet(layoutGid)
        .values({ startRowIndex: 9, endRowIndex: 15, endColumnIndex: 2 }),
    ).toEqual([
      ["c:lyb:entry", "c:lyb:amount"],
      [null, null],
      [null, null],
      ["Entry", "Amount"],
      ["Below one", 100],
      ["Below two", 200],
    ]);
  });

  it("pulls the Table below up intact when a row above it is deleted", () => {
    const { ss, grid } = fetchedLayout();
    ss.table("layoutLeft").row(1).delete();
    ss.batchUpdateGSheets();

    expect(placedRange(grid, "layoutLeft")).toEqual([3, 6, 0, 2]);
    expect(placedRange(grid, "layoutBelow")).toEqual([10, 13, 0, 2]);
    expect(placedRange(grid, "layoutRight")).toEqual([3, 6, 3, 5]);
    expect(
      grid
        .sheet(layoutGid)
        .values({ startRowIndex: 4, endRowIndex: 6, endColumnIndex: 5 }),
    ).toEqual([
      ["Left one", 1, null, "Right one", 10],
      ["Left three", 3, null, "Right two", 20],
    ]);
  });

  it("shifts the Table beside it right, head rows included, when a column is inserted at its end", () => {
    const { ss, grid } = fetchedLayout();
    ss.table("layoutLeft").raw.appendColumn({
      columnId: "c:lyl:note",
      header: "Note",
    });
    ss.batchUpdateGSheets();

    expect(placedRange(grid, "layoutLeft")).toEqual([3, 7, 0, 3]);
    expect(placedRange(grid, "layoutRight")).toEqual([3, 6, 4, 6]);
    expect(placedRange(grid, "layoutBelow")).toEqual([11, 14, 0, 2]);
    expect(
      grid
        .sheet(layoutGid)
        .values({ startRowIndex: 0, endRowIndex: 1, endColumnIndex: 6 }),
    ).toEqual([
      [
        "c:lyl:entry",
        "c:lyl:amount",
        "c:lyl:note",
        null,
        "c:lyr:entry",
        "c:lyr:amount",
      ],
    ]);
  });

  it("keeps reading the Table below once growth has pushed it down, in the same run", () => {
    const { ss } = fetchedLayout();
    ss.table("layoutLeft").appendRowWithVals({ entry: "Left four", amount: 4 });
    ss.batchUpdateGSheets();
    ss.table("layoutBelow").prepFetchColumnsFull("entry", "amount");
    ss.fetchAllPrepped();

    expect(
      ss.table("layoutBelow").rows.map((row) => row.value("entry")),
    ).toEqual(["Below one", "Below two"]);
  });

  it("keeps reading the Table beside once a column insert has shifted it, in the same run", () => {
    const { ss } = fetchedLayout();
    ss.table("layoutLeft").raw.appendColumn({
      columnId: "c:lyl:note",
      header: "Note",
    });
    ss.batchUpdateGSheets();
    ss.table("layoutRight").prepFetchColumnsFull("entry", "amount");
    ss.fetchAllPrepped();

    expect(
      ss.table("layoutRight").rows.map((row) => row.value("amount")),
    ).toEqual([10, 20]);
  });

  describe("refuses an edit protection on a Table whose properties have not arrived", () => {
    function fetchedLeftOnly(): SpreadsheetNamed {
      stubSheetsService({ sheets: [layoutSheet()] });
      const ss = SpreadsheetNamed.init();
      ss.table("layoutLeft").column("amount").prepFetchSpecific([0, 1]);
      ss.fetchAllPrepped();
      return ss;
    }

    it("behind growth that would push it down", () => {
      const ss = fetchedLeftOnly();
      ss.table("layoutLeft").appendRowWithVals({
        entry: "Left four",
        amount: 4,
      });

      expect(() =>
        ss.table("layoutBelow").raw.row(0).cell(1).addEditWarning(),
      ).toThrowError(/Table is unknown for sheetGid \d+/);
    });

    it("behind a column insert that would push it along", () => {
      const ss = fetchedLeftOnly();
      ss.table("layoutLeft").raw.appendColumn({
        columnId: "c:lyl:note",
        header: "Note",
      });

      expect(() =>
        ss.table("layoutRight").raw.row(0).cell(1).addEditLock(),
      ).toThrowError(/Table is unknown for sheetGid \d+/);
    });
  });
});
