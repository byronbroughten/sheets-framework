import type { Chore } from "../../src/chores/Chore";
import { DevFixtureBuilder } from "./devFixtures/DevFixtureBuilder";
import { devSpreadsheetId } from "./devFixtures/devFixtureSheets";

export const buildDevFixtures: Chore = {
  description:
    "Dev spreadsheet only: creates any missing fixture tab (Item, Value Types, Log, Run Item, Computed, Dates) with its Table, column types, column IDs and rows, then ticks Let api access and declares the Empty value allowed exemplars. Needs the config floor from dev:gen:configs; run dev:gen:configs again afterwards for the dev generated/.",
  action: (_ss, { spreadsheetId }) => {
    if (spreadsheetId !== devSpreadsheetId) {
      throw new Error(
        `buildDevFixtures refused: it runs only on the dev spreadsheet (${devSpreadsheetId}), not ${spreadsheetId}.`,
      );
    }
    return DevFixtureBuilder.init().ensureFixtures();
  },
};
