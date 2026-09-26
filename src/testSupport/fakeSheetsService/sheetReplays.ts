import { type FakeSpreadsheet, fakeSpreadsheet } from "./fakeSpreadsheet";

type Response = GoogleAppsScript.Sheets.Schema.Response;

// Google's size for a tab added with no grid properties.
const defaultGrid = { rowCount: 1000, columnCount: 26 } as const;

export const sheetReplays = {
  addSheet(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.AddSheetRequest,
  ): Response {
    const properties = request.properties ?? {};
    const sheetId =
      properties.sheetId ??
      Math.max(0, ...spreadsheet.sheets.map((sheet) => sheet.sheetId)) + 1;
    const title = properties.title ?? `Sheet${spreadsheet.sheets.length + 1}`;
    spreadsheet.sheets.forEach((sheet) => {
      if (sheet.sheetId === sheetId || sheet.title === title) {
        throw new Error(
          `addSheet: a sheet with gid ${sheetId} or title "${title}" already exists.`,
        );
      }
    });
    const rowCount =
      properties.gridProperties?.rowCount ?? defaultGrid.rowCount;
    const columnCount =
      properties.gridProperties?.columnCount ?? defaultGrid.columnCount;
    spreadsheet.sheets.splice(
      properties.index ?? spreadsheet.sheets.length,
      0,
      {
        sheetId,
        title,
        rows: [],
        extraTables: [],
        rowCount,
        columnCount,
        hiddenRowIndexes: [],
        hiddenColumnIndexes: [],
      },
    );
    return {
      addSheet: {
        properties: {
          sheetId,
          title,
          gridProperties: { rowCount, columnCount },
        },
      },
    };
  },
  deleteSheet(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.DeleteSheetRequest,
  ): Response {
    const sheet = fakeSpreadsheet.sheet(spreadsheet, request.sheetId);
    spreadsheet.sheets.splice(spreadsheet.sheets.indexOf(sheet), 1);
    return {};
  },
  updateSheetProperties(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.UpdateSheetPropertiesRequest,
  ): Response {
    const properties = request.properties ?? {};
    const sheet = fakeSpreadsheet.sheet(spreadsheet, properties.sheetId);
    (request.fields ?? "").split(",").forEach((rawField) => {
      const field = rawField.trim();
      if (field === "title") {
        sheet.title = properties.title ?? "";
      } else if (field === "gridProperties.rowCount") {
        sheet.rowCount = properties.gridProperties?.rowCount ?? sheet.rowCount;
      } else if (field === "gridProperties.columnCount") {
        sheet.columnCount =
          properties.gridProperties?.columnCount ?? sheet.columnCount;
      } else {
        throw new Error(
          `The fake Sheets service does not replay updateSheetProperties field "${field}".`,
        );
      }
    });
    return {};
  },
};
