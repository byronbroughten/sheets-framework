import type { SpreadsheetNamed } from "../../src/04_SpreadsheetNamed/SpreadsheetNamed";
import type { Chore } from "../../src/chores/Chore";
import { devSpreadsheetId } from "./devFixtures/devFixtureSheets";

const layoutTableNames = ["layoutLeft", "layoutRight"] as const;
const noteColumnId = "c:lyl:note";

export const exerciseLayout: Chore = {
  description:
    "Dev spreadsheet only: reads both Layout Tables and the Layout grid's column count. On a Layout tab built to its recipe, it then sends one batch that appends a row to layoutLeft, inserts a column at its end (pushing layoutRight right) and deletes its second row; run it again to read the pushed layoutRight without regenerating. The batch leaves the tab off its recipe: delete the Layout tab, then rerun buildDevFixtures and dev:gen:configs.",
  action: (ss, { spreadsheetId }) => {
    if (spreadsheetId !== devSpreadsheetId) {
      throw new Error(
        `exerciseLayout refused: it runs only on the dev spreadsheet (${devSpreadsheetId}), not ${spreadsheetId}.`,
      );
    }
    const report = readLayout(ss);
    const left = ss.table("layoutLeft");
    if (left.raw.columnResolver.hasColumnId(noteColumnId)) {
      return `${report}\nAlready exercised; nothing sent.`;
    }
    left.appendRowWithVals({ entry: "Left four", amount: 4 });
    left.raw.appendColumn({ columnId: noteColumnId, header: "Note" });
    left.rowsFiltered({ entry: "Left two" }).forEach((row) => row.delete());
    ss.batchUpdateGSheets();
    return report;
  },
};

function readLayout(ss: SpreadsheetNamed): string {
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
  const gridColumnCount = ss.table("layoutLeft").raw.sheet.columnCount;
  return ["Read:", ...lines, `  Layout grid: ${gridColumnCount} columns`].join(
    "\n",
  );
}
