import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { type PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { SpreadsheetRaw } from "../02_SpreadsheetRaw/SpreadsheetRaw";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  type FakeSheetsService,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { SyncCoordinator } from "./SyncCoordinator";
import {
  emittedItemColumn,
  itemAtEachFlush,
  itemColumns,
  itemGid,
  itemRowId,
  itemSheet,
  noteColIndex,
  runItem,
  runWithNoAlterations,
  seedFixture,
} from "./syncCoordinatorTestSupport";

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

const compiledNote = {
  isFormula: true,
  value: `=SINGLE(${getTableTraitByName("item", "tableName")}[${itemColumns.id.header}])`,
};

function noteCellsOf(service: FakeSheetsService): unknown[] {
  return service.grid
    .sheet(itemGid)
    .bodyRows()
    .map((row) => row[noteColIndex]);
}

describe("SyncCoordinator.run formulas", () => {
  const formulaAlteration: PendingAlterations = {
    tables: { item: { columns: { optionalNote: { formula: "=${@id}" } } } },
  };

  it("fills every body row with the compiled formula, and the configs record it in brace form", () => {
    const service = seedFixture({
      item: {
        // The name the previous configs record, which the compiled reference uses.
        name: getTableTraitByName("item", "tableName"),
        bodyRows: [{ id: "r:itm:one" }, { id: "r:itm:two" }],
      },
    });

    const { files } = runItem({
      isSending: true,
      alterations: formulaAlteration,
    });

    expect(noteCellsOf(service)).toEqual([compiledNote, compiledNote]);
    expect(
      emittedItemColumn(Val.assert(files, "emitted files"), "optionalNote"),
    ).toContain('"formula": "=${@id}"');
  });

  it("gives a row appended afterward through Raw the compiled formula", () => {
    const sheet = itemSheet({
      name: getTableTraitByName("item", "tableName"),
      bodyRows: [{ id: "r:itm:one" }],
    });
    // Room below the Table, so the append grows it through `PASTE_FORMULA` and not the grid's edge.
    const service = stubSheetsService({
      sheets: [{ ...sheet, rows: [...Val.assert(sheet.rows, "rows"), [], []] }],
    });
    runItem({ isSending: true, alterations: formulaAlteration });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table("item").appendDataRow([noteColIndex]);
    raw.batchUpdateGSheets();

    expect(noteCellsOf(service)).toEqual([compiledNote, compiledNote]);
  });

  it("reads the column before a send's flush, holding the cell a person entered", () => {
    const [atFirstFlush] = itemAtEachFlush({
      isSending: true,
      alterations: formulaAlteration,
    });

    expect(atFirstFlush).toEqual({
      colIndexesRead: [noteColIndex],
      note: "kept",
    });
  });

  it("refuses a brace naming no column before any read", () => {
    const service = seedFixture();

    const { result } = runItem({
      isSending: true,
      alterations: {
        tables: {
          item: { columns: { optionalNote: { formula: "=${@nope}" } } },
        },
      },
    });

    expect(result.refusals).toEqual([
      expect.objectContaining({ kind: "unknownBrace", brace: "${@nope}" }),
    ]);
    expect(service.getCalls).toEqual([]);
  });
});

describe("SyncCoordinator formula readings", () => {
  const idHeader = itemColumns.id.header;
  const configsName = getTableTraitByName("item", "tableName");

  function liveNoteFormula(): unknown {
    const coordinator = SyncCoordinator.init();
    runWithNoAlterations(coordinator);
    return coordinator.columnConfigOperator.traitReadings().live.tables.item
      ?.columns.optionalNote?.formula;
  }

  // Item's note column holding a formula that names Item by the given name.
  function seedNaming(itemTableName: string): void {
    seedFixture({
      item: {
        name: itemTableName,
        bodyRows: [
          {
            id: itemRowId,
            optionalNote: {
              value: itemRowId,
              isFormula: true,
              formula: `=SINGLE(${itemTableName}[${idHeader}])`,
            },
          },
        ],
      },
    });
  }

  it("reads a formula naming a renamed Table as unchanged, in brace form", () => {
    seedNaming(configsName);
    const underConfigsName = liveNoteFormula();
    seedNaming("Stock");

    expect(underConfigsName).toBe("=${@id}");
    expect(liveNoteFormula()).toBe("=${@id}");
  });

  it("reads a formula sent beside a Table rename as the alteration's brace form", () => {
    const service = seedFixture({ item: { name: configsName } });

    const { result } = runItem({
      isSending: true,
      alterations: {
        tables: {
          item: {
            name: "Stock",
            columns: { optionalNote: { formula: "=${@id}" } },
          },
        },
      },
    });

    expect(result.outcome.kind).toBe("sent");
    expect(noteCellsOf(service)).toEqual([
      { isFormula: true, value: `=SINGLE(Stock[${idHeader}])` },
    ]);
    expect(liveNoteFormula()).toBe("=${@id}");
  });
});
