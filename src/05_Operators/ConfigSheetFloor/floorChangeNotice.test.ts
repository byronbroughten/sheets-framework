import { describe, expect, it } from "vitest";

import { configSheetFloorSeed } from "../../01_SpreadsheetSchema/configSheetFloorSeed";
import { getSheetTraitByName } from "../../01_SpreadsheetSchema/sheetConfigsTypes";
import { floorChangeNotice } from "./floorChangeNotice";

const businessSheetGid = 9001;
const spreadsheetConfigGid = getSheetTraitByName(
  "spreadsheetConfig",
  "sheetGid",
);
const sheetConfigGid = getSheetTraitByName("sheetConfig", "sheetGid");
const columnConfigGid = getSheetTraitByName("columnConfig", "sheetGid");
const valueConfigGid = getSheetTraitByName("valueConfig", "sheetGid");

function liveTitles(
  overrides: Record<number, string | null> = {},
): Map<number, string> {
  const titles = new Map<number, string | null>([
    [businessSheetGid, "Widget"],
    [spreadsheetConfigGid, configSheetFloorSeed.spreadsheetConfig.title],
    [sheetConfigGid, configSheetFloorSeed.sheetConfig.title],
    [columnConfigGid, configSheetFloorSeed.columnConfig.title],
    [valueConfigGid, configSheetFloorSeed.valueConfig.title],
  ]);
  Object.entries(overrides).forEach(([sheetGid, title]) => {
    titles.set(Number(sheetGid), title);
  });
  return new Map(
    [...titles].flatMap(([sheetGid, title]) =>
      title === null ? [] : [[sheetGid, title] as const],
    ),
  );
}

describe("floorChangeNotice", () => {
  it("says a renamed Value Config is managed and will switch back on the next config sync, and closes by itself", () => {
    const notice = floorChangeNotice(
      "other",
      liveTitles({ [valueConfigGid]: "Values" }),
    );
    expect(notice).toEqual({
      title: "Value Config is managed",
      message:
        'This tab keeps the name "Value Config". Your rename will switch back the next time configs sync.',
      untilClosed: false,
    });
  });

  it("tells whoever deleted Value Config to press Undo now, because the next sync recreates it empty, and stays until closed", () => {
    const notice = floorChangeNotice(
      "sheetRemoved",
      liveTitles({ [valueConfigGid]: null }),
    );
    expect(notice).toEqual({
      title: "Value Config was deleted",
      message:
        "Press Undo (Ctrl+Z, or ⌘Z on a Mac) now to get it back with its data. If you don't, the next sync recreates it empty.",
      untilClosed: true,
    });
  });

  it("says nothing when a business tab is renamed", () => {
    expect(
      floorChangeNotice("other", liveTitles({ [businessSheetGid]: "Gadgets" })),
    ).toBeUndefined();
  });

  it.each([
    ["Spreadsheet Config", spreadsheetConfigGid],
    ["Sheet Config", sheetConfigGid],
    ["Column Config", columnConfigGid],
  ])("says nothing when warned %s is renamed", (title, sheetGid) => {
    expect(
      floorChangeNotice("other", liveTitles({ [sheetGid]: `Old ${title}` })),
    ).toBeUndefined();
  });

  it("says nothing when a business tab is deleted", () => {
    expect(
      floorChangeNotice(
        "sheetRemoved",
        liveTitles({ [businessSheetGid]: null }),
      ),
    ).toBeUndefined();
  });

  it("says nothing for a change that renames or deletes nothing", () => {
    expect(floorChangeNotice("other", liveTitles())).toBeUndefined();
    expect(floorChangeNotice("sheetRemoved", liveTitles())).toBeUndefined();
  });
});
