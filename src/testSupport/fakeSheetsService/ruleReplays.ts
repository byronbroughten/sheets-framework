import { type FakeSpreadsheet, fakeSpreadsheet } from "./fakeSpreadsheet";

type Response = GoogleAppsScript.Sheets.Schema.Response;

// Conditional formats and edit protections: an add inserts, a delete renumbers.
export const ruleReplays = {
  addConditionalFormatRule(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.AddConditionalFormatRuleRequest,
  ): Response {
    const rule = request.rule ?? {};
    const sheet = fakeSpreadsheet.sheet(spreadsheet, rule.ranges?.[0]?.sheetId);
    sheet.conditionalFormats ??= [];
    sheet.conditionalFormats.splice(request.index ?? 0, 0, rule);
    return {};
  },
  deleteConditionalFormatRule(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.DeleteConditionalFormatRuleRequest,
  ): Response {
    const sheet = fakeSpreadsheet.sheet(spreadsheet, request.sheetId);
    const rules = sheet.conditionalFormats ?? [];
    const index = request.index ?? 0;
    if (index >= rules.length) {
      throw new Error(
        `deleteConditionalFormatRule: sheet ${sheet.sheetId} has no rule at index ${index}.`,
      );
    }
    rules.splice(index, 1);
    return {};
  },
  addProtectedRange(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.AddProtectedRangeRequest,
  ): Response {
    const add = request.protectedRange ?? {};
    const sheet = fakeSpreadsheet.sheet(spreadsheet, add.range?.sheetId);
    spreadsheet.lastProtectedRangeId += 1;
    const stored = {
      ...add,
      protectedRangeId: spreadsheet.lastProtectedRangeId,
    };
    sheet.protectedRanges ??= [];
    sheet.protectedRanges.push(stored);
    return { addProtectedRange: { protectedRange: stored } };
  },
  deleteProtectedRange(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.DeleteProtectedRangeRequest,
  ): Response {
    const removeId = request.protectedRangeId;
    spreadsheet.sheets.forEach((sheet) => {
      if (sheet.protectedRanges === undefined) return;
      sheet.protectedRanges = sheet.protectedRanges.filter(
        (protection) => protection.protectedRangeId !== removeId,
      );
    });
    return {};
  },
};
