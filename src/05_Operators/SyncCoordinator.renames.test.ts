import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  type FakeSheetsService,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { parseIdNameMapFileSource } from "./parseIdNameMapFileSource";
import {
  emittedItemColumn,
  itemColumns,
  itemGid,
  itemSheet,
  runItem,
  seedFixture,
} from "./syncCoordinatorTestSupport";

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("SyncCoordinator.run renames", () => {
  // A typed column on every Item column, so a test sees the types survive the rename.
  const itemTypes = { 0: "TEXT", 1: "DOUBLE", 2: "TEXT", 3: "TEXT" };
  const optionalNoteIndex = Object.keys(itemColumns).indexOf("optionalNote");
  const emptied = [
    "// sync rewrites this file: each alteration that lands is removed, and comments here won't survive a run.",
    'import { makePendingAlterations } from "@byronbroughten/sheets-framework";',
    "",
    "export const pendingAlterations = makePendingAlterations({});",
    "",
  ].join("\n");

  function liveItemTable(
    service: FakeSheetsService,
  ): GoogleAppsScript.Sheets.Schema.Table {
    return Val.assert(service.grid.sheet(itemGid).tables[0], "live Item Table");
  }

  it("renames a column by writing its header cell, keeping its column type, and the configs follow the new header", () => {
    const service = seedFixture({ item: { columnTypes: itemTypes } });

    const { result, files } = runItem({
      isSending: true,
      alterations: {
        tables: { item: { columns: { optionalNote: { header: "Remark" } } } },
      },
    });

    const emitted = Val.assert(files, "emitted files");
    expect(
      liveItemTable(service).columnProperties?.[optionalNoteIndex],
    ).toMatchObject({ columnName: "Remark", columnType: "TEXT" });
    expect(liveItemTable(service).columnProperties?.[1]).toMatchObject({
      columnType: "DOUBLE",
    });
    expect(emittedItemColumn(emitted, "remark")).toContain(
      `"columnId": "${itemColumns.optionalNote.columnId}", "header": "Remark"`,
    );
    expect(emitted.columnConfigs).not.toContain('"optionalNote"');
    expect(
      parseIdNameMapFileSource(emitted.idNameMap).tables.item?.columns,
    ).toMatchObject({ remark: itemColumns.optionalNote.columnId });
    expect(emitted.alterations).toBe(emptied);
    expect(result).toMatchObject({
      mode: "send",
      outcome: { kind: "sent", summary: "1 alteration landed" },
      landed: [
        {
          kind: "send",
          subject: { table: "item", column: "optionalNote", path: "header" },
          what: "set to Remark",
        },
      ],
      didntLand: [],
    });
  });

  it("renames a Table, keeping its column types, and the configs key it by its new name", () => {
    // The name the previous configs record, so the rename isn't a person's change.
    const service = seedFixture({
      item: {
        columnTypes: itemTypes,
        name: getTableTraitByName("item", "tableName"),
      },
    });

    const { result, files } = runItem({
      isSending: true,
      alterations: { tables: { item: { name: "Stock" } } },
    });

    const emitted = Val.assert(files, "emitted files");
    expect(liveItemTable(service)).toMatchObject({ name: "Stock" });
    expect(liveItemTable(service).columnProperties?.[1]).toMatchObject({
      columnType: "DOUBLE",
    });
    expect(emitted.tableConfigs).toContain(
      '"stock": { "tableId": "item", "tableName": "Stock"',
    );
    expect(
      parseIdNameMapFileSource(emitted.idNameMap).tables.item,
    ).toMatchObject({ name: "stock" });
    expect(emitted.alterations).toBe(emptied);
    expect(result).toMatchObject({
      outcome: { kind: "sent", summary: "1 alteration landed" },
      landed: [
        {
          kind: "send",
          subject: { table: "item", path: "name" },
          what: "set to Stock",
        },
      ],
      didntLand: [],
    });
  });

  it("reports a trait Sheets kept another value for under Didn't land, the configs taking that value and the entry staying pending", () => {
    const service = stubSheetsService({
      sheets: [
        itemSheet({
          columnTypes: itemTypes,
          name: getTableTraitByName("item", "tableName"),
        }),
      ],
      afterEachBatchUpdate(spreadsheet) {
        const table = Val.assert(spreadsheet.sheets[0]?.tables[0], "Item");
        table.name = "Item";
      },
    });

    const { result, files } = runItem({
      isSending: true,
      alterations: {
        tables: {
          item: {
            name: "Stock",
            columns: { optionalNote: { header: "Remark" } },
          },
        },
      },
    });

    const emitted = Val.assert(files, "emitted files");
    expect(liveItemTable(service)).toMatchObject({ name: "Item" });
    expect(
      liveItemTable(service).columnProperties?.[optionalNoteIndex],
    ).toMatchObject({ columnName: "Remark" });
    expect(emitted.tableConfigs).toContain(
      '"item": { "tableId": "item", "tableName": "Item"',
    );
    expect(emittedItemColumn(emitted, "remark")).toContain(
      '"header": "Remark"',
    );
    expect(emitted.alterations).toBe(
      [
        "// sync rewrites this file: each alteration that lands is removed, and comments here won't survive a run.",
        'import { makePendingAlterations } from "@byronbroughten/sheets-framework";',
        "",
        "export const pendingAlterations = makePendingAlterations({",
        '  "tables": {',
        '    "item": { "name": "Stock" }',
        "  }",
        "});",
        "",
      ].join("\n"),
    );
    expect(result).toMatchObject({
      mode: "send",
      outcome: {
        kind: "sent",
        summary: "1 alteration landed; 1 didn't land",
      },
      landed: [
        {
          kind: "send",
          subject: { table: "item", column: "optionalNote", path: "header" },
          what: "set to Remark",
        },
      ],
      didntLand: [
        {
          subject: { table: "item", path: "name" },
          what: "set to Stock; Sheets kept Item",
        },
      ],
    });
  });

  it("sends neither rename on a dry run", () => {
    const service = seedFixture({ item: { columnTypes: itemTypes } });

    runItem({
      isSending: false,
      alterations: {
        tables: {
          item: {
            name: "Stock",
            columns: { optionalNote: { header: "Remark" } },
          },
        },
      },
    });

    expect(service.batchUpdateCount()).toBe(0);
    expect(liveItemTable(service)).toMatchObject({ name: "Item" });
    expect(
      liveItemTable(service).columnProperties?.[optionalNoteIndex],
    ).toMatchObject({ columnName: itemColumns.optionalNote.header });
  });
});
