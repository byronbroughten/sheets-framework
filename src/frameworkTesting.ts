// The framework's test entry (`./testing`), for `*.test.ts` files and the test setup file only.
export { installConfigs } from "./01_SpreadsheetSchema/configReaders/configRegister.ts";
export { EndpointRun } from "./06_API/EndpointRun.ts";
export { stubLogger } from "./testSupport/fakeAppsScriptGlobals.ts";
export {
  buildGridRows,
  stubSheetsService,
  type FakeCell,
  type FakeCellValue,
  type FakeSheetProperties,
  type FakeSheetsService,
} from "./testSupport/fakeSheetsService.ts";
export {
  type FakeBodyRow,
  type FakeTableColumn,
  type FakeTableSheetProps,
  fakeTableSheet,
} from "./testSupport/fakeSheetsService/fakeTableSheet.ts";
export type {
  FakeGridRange,
  FakeGridView,
  FakeSheetView,
} from "./testSupport/fakeSheetsService/gridView.ts";
