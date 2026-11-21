import type { Chore } from "../../src/chores/Chore";
import { DevFixtureBuilder } from "./devFixtures/DevFixtureBuilder";
import { devSpreadsheetId } from "./devFixtures/devFixtureSheets";

export const buildDevFixtures: Chore = {
  description:
    "Dev spreadsheet only: creates any missing fixture tab (Item, Value Types, Log, Run Item, Computed, Dates, Layout) with its Tables, column types, column IDs and rows, each Table under its fixed Table ID. Run dev:sync afterwards for the dev generated/.",
  action: (_ss, { spreadsheetId }) => {
    if (spreadsheetId !== devSpreadsheetId) {
      throw new Error(
        `buildDevFixtures refused: it runs only on the dev spreadsheet (${devSpreadsheetId}), not ${spreadsheetId}.`,
      );
    }
    return DevFixtureBuilder.init().ensureFixtures();
  },
};
