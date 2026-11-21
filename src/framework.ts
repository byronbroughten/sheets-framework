// eslint-disable-next-line @typescript-eslint/triple-slash-reference -- ambient declarations ship with the entry, not as a module
/// <reference path="./TypeDeclarations/google-apps-script-ext.d.ts" />
// The framework's public runtime entry (`.`): app code imports nothing else, bar `./testing` in tests.
export type { Register } from "./01_SpreadsheetSchema/configReaders/configRegister.ts";
export {
  makeColumnConfigs,
  makeConditionalFormatConfigs,
  makeRetiredIdPrefixes,
  makeTabConfigs,
  makeTableConfigs,
  makeValueConfigs,
} from "./01_SpreadsheetSchema/configReaders/makeConfigs.ts";
export {
  makePendingAlterations,
  type PendingAlterations,
} from "./01_SpreadsheetSchema/configReaders/intendedState.ts";
export { TableBaseNamed } from "./04_SpreadsheetNamed/ClassBases/TableBaseNamed.ts";
export {
  SpreadsheetBaseNamed,
  type SpreadsheetNamedProps,
} from "./04_SpreadsheetNamed/ClassBases/SpreadsheetBaseNamed.ts";
export { RowNamed } from "./04_SpreadsheetNamed/RowNamed.ts";
export { TableNamed } from "./04_SpreadsheetNamed/TableNamed.ts";
export { SpreadsheetNamed } from "./04_SpreadsheetNamed/SpreadsheetNamed.ts";
export type { RowIdByName } from "./04_SpreadsheetNamed/Types/RowIdByName.ts";
export type {
  ActionReturn,
  Endpoint,
  Endpoints,
  RowReports,
  RunReport,
} from "./06_API/Endpoints.ts";
export { AppsScriptApi as Api } from "./appsScriptHost/AppsScriptApi.ts";
export type { Chore } from "./chores/Chore.ts";
export type {
  DateInRange,
  DateRange,
  FirstAndLastOfMonth,
  MonthRange,
  MonthYear,
  MonthYearRange,
  SerialDate,
  Ymd,
} from "@byronbroughten/utils/serial-date";
