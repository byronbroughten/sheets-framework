// The scripts project installs no configs, and valueTypesTableSheet reads them as it loads.
import "../dev/installDevConfigs.ts";

import { describe, expect, it } from "vitest";

import { fakeSheetsHttpTransport } from "../src/testSupport/fakeSheetsHttpTransport.ts";
import { stubSheetsService } from "../src/testSupport/fakeSheetsService.ts";
import {
  filledValueTypesRow,
  valueTypesGid,
  valueTypesTableSheet,
} from "../src/testSupport/valueTypesTableSheet.ts";
import { choreSheetsConfig, runChoreOver } from "./choreFixture.ts";

// A chore that reads a column, then writes only into blank cells, which the gate lets through.
const appendNoteChoreSource = `
export const appendNote = {
  description: "appends a Note column to Value Types",
  action(ss) {
    const table = ss.table("valueTypes");
    table.prepFetchColumnsFull("stringValue");
    ss.fetchAllPrepped();
    table.raw.appendColumn({ columnId: "c:vty:note", header: "Note" });
    ss.batchUpdateGSheets();
  },
};
`;

const sheetsConfig = choreSheetsConfig({
  name: "appendNote",
  source: appendNoteChoreSource,
});

describe("a chore's --send over the fake Sheets transport", () => {
  it("leaves its write on the fake spreadsheet's grid", async () => {
    const fake = stubSheetsService({
      sheets: [valueTypesTableSheet([filledValueTypesRow])],
    });
    const columnCount = fake.grid.sheet(valueTypesGid).columnCount;

    const sent = await runChoreOver({
      sheetsConfig,
      transport: fakeSheetsHttpTransport(fake),
      argv: ["appendNote", "--send"],
    });

    const sheet = fake.grid.sheet(valueTypesGid);
    expect(sent.exitCode).toBeUndefined();
    expect(sheet.columnCount).toBe(columnCount + 1);
    expect(sheet.values().map((row) => row[columnCount])).toEqual(
      expect.arrayContaining(["c:vty:note", "Note"]),
    );
  });
});
