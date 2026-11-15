import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { type PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { SpreadsheetRaw } from "../02_SpreadsheetRaw/SpreadsheetRaw";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import { stubSheetsService } from "../testSupport/fakeSheetsService";
import {
  emittedItemColumn,
  itemAtEachFlush,
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

describe("SyncCoordinator.run formulas", () => {
  const optionalNoteIndex = Object.keys(itemColumns).indexOf("optionalNote");
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

    const noteCells = service.grid
      .sheet(itemGid)
      .bodyRows()
      .map((row) => row[optionalNoteIndex]);
    const compiled = {
      isFormula: true,
      value: `=SINGLE(${getTableTraitByName("item", "tableName")}[${itemColumns.id.header}])`,
    };
    expect(noteCells).toEqual([compiled, compiled]);
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
    raw.table("item").appendDataRow([optionalNoteIndex]);
    raw.batchUpdateGSheets();

    const noteCells = service.grid
      .sheet(itemGid)
      .bodyRows()
      .map((row) => row[optionalNoteIndex]);
    const compiled = {
      isFormula: true,
      value: `=SINGLE(${getTableTraitByName("item", "tableName")}[${itemColumns.id.header}])`,
    };
    expect(noteCells).toEqual([compiled, compiled]);
  });

  describe("a body cell of the column being given a formula", () => {
    it("holds the value a person entered when the send's flush goes out", () => {
      const [atFirstFlush] = itemAtEachFlush({
        isSending: true,
        alterations: formulaAlteration,
      });

      expect(atFirstFlush?.note).toBe("kept");
    });

    it("is read before the send's flush, and no other column is", () => {
      const [atFormulaFlush] = itemAtEachFlush({
        isSending: true,
        alterations: formulaAlteration,
      });
      const [atBareFlush] = itemAtEachFlush({
        isSending: true,
        alterations: { tables: { item: { columns: {} } } },
      });

      expect(atFormulaFlush?.colIndexesRead).toEqual([optionalNoteIndex]);
      expect(atBareFlush?.colIndexesRead ?? []).toEqual([]);
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
