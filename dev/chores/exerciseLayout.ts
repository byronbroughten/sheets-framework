import type { SpreadsheetNamed } from "../../src/04_SpreadsheetNamed/SpreadsheetNamed";
import type { Chore } from "../../src/chores/Chore";
import { devSpreadsheetId } from "./devFixtures/devFixtureSheets";

const layoutTableNames = ["layoutLeft", "layoutRight", "layoutBelow"] as const;

export const exerciseLayout: Chore = {
  description:
    "Dev spreadsheet only: reads the three Layout Tables, then sends one batch that appends a row to layoutLeft (pushing layoutBelow down), inserts a column at its end (shifting layoutRight right) and deletes its second row (pulling layoutBelow back up). The batch leaves the tab off its recipe: delete the Layout tab, then rerun buildDevFixtures and dev:gen:configs.",
  action: (ss, { spreadsheetId }) => {
    if (spreadsheetId !== devSpreadsheetId) {
      throw new Error(
        `exerciseLayout refused: it runs only on the dev spreadsheet (${devSpreadsheetId}), not ${spreadsheetId}.`,
      );
    }
    const report = readEveryTable(ss);
    const left = ss.table("layoutLeft");
    left.appendRowWithVals({ entry: "Left four", amount: 4 });
    left.raw.appendColumn({ columnId: "c:lyl:note", header: "Note" });
    left.rowsFiltered({ entry: "Left two" }).forEach((row) => row.delete());
    ss.batchUpdateGSheets();
    return report;
  },
};

function readEveryTable(ss: SpreadsheetNamed): string {
  layoutTableNames.forEach((tableName) =>
    ss.table(tableName).prepFetchColumnsFull("entry", "amount"),
  );
  ss.fetchAllPrepped();
  const lines = layoutTableNames.map((tableName) => {
    const table = ss.table(tableName);
    const { headerRowIndex, startColIndex } = table.raw.tableOrigin();
    const rows = table.rows
      .map((row) => `${row.value("entry")} ${row.value("amount")}`)
      .join(", ");
    return `  ${tableName} at ${ss.schema.positionLabel(headerRowIndex, startColIndex)}, ${table.raw.columnCount} columns: ${rows}`;
  });
  return ["Read:", ...lines].join("\n");
}
