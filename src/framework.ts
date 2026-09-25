// eslint-disable-next-line @typescript-eslint/triple-slash-reference -- ambient declarations ship with the entry, not as a module
/// <reference path="./TypeDeclarations/google-apps-script-ext.d.ts" />
// The framework's public runtime entry (`.`): app code imports nothing else, bar `./testing` in tests.
export type { Register } from "./01_SpreadsheetSchema/configRegister";
export { SheetBaseNamed } from "./04_SpreadsheetNamed/ClassBases/SheetBaseNamed";
export {
  SpreadsheetBaseNamed,
  type SpreadsheetNamedProps,
} from "./04_SpreadsheetNamed/ClassBases/SpreadsheetBaseNamed";
export { RowNamed } from "./04_SpreadsheetNamed/RowNamed";
export { SheetNamed } from "./04_SpreadsheetNamed/SheetNamed";
export { SpreadsheetNamed } from "./04_SpreadsheetNamed/SpreadsheetNamed";
export type { RowIdByName } from "./04_SpreadsheetNamed/Types/RowIdByName";
export type {
  ActionReturn,
  Endpoint,
  Endpoints,
  RowReports,
  RunReport,
} from "./06_API/Endpoints";
export { AppsScriptApi as Api } from "./appsScriptHost/AppsScriptApi";
export type { Chore } from "./chores/Chore";
export {
  SerialDate,
  type DateInRange,
  type DateRange,
  type FirstAndLastOfMonth,
  type MonthRange,
  type MonthYear,
  type MonthYearRange,
  type Ymd,
} from "./utils/SerialDate";
