import type {
  ColumnFullName,
  ColumnNameFiltered,
  SheetNameOf,
} from "../01_SpreadsheetSchema/columnConfigsTypes";
import type { FloorTabName } from "../01_SpreadsheetSchema/configSheetFloorSeed";
import type { SheetNameSimple } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import type { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import type { CheckboxColumnName } from "../05_Operators/CheckboxColumnOperator";

// The only columns a run can report into: writable, and holding a sentence.
export type FeedbackColumnName<SN extends SheetNameSimple> = ColumnNameFiltered<
  SN,
  "string",
  false
>;

export interface EndpointArgs {
  selectedRowIndexes: number[];
  isChecked: boolean;
}

// Every state but running, which the framework writes for itself.
export type RunStateReported = "success" | "warning" | "failure";

// Neither warning nor failure has a fallback sentence, so the type demands one.
export type RunReport =
  | { runState?: "success"; message?: string }
  | { runState: Exclude<RunStateReported, "success">; message: string };

// Keyed by the same base-zero grid index the action is handed as its rows.
export type RowReports = Map<number, RunReport>;

// `void`, so an action with nothing to report needs no explicit `undefined`.
export type ActionReturn = void | string | (RunReport & { rows?: RowReports });

export type EndpointAction = (
  ss: SpreadsheetNamed,
  args: EndpointArgs,
) => ActionReturn;

export interface Endpoint<SN extends SheetNameSimple> {
  action: EndpointAction;
  timeLastRan?: FeedbackColumnName<SN>;
  runStatus?: FeedbackColumnName<SN>;
  // Inline, not a named type: a named one compares by variance, which the widened dispatch boundary rejects.
  selector?: {
    column: CheckboxColumnName<SN>;
    retainSelection?: boolean;
    requireOneRow?: boolean;
  };
  runOnUncheck?: boolean;
}

// The entry as the dispatch hands it over — a structural copy, for the same reason.
export type EndpointDispatched<SN extends SheetNameSimple> = {
  [K in keyof Endpoint<SN>]: Endpoint<SN>[K];
};

// Each key carries its own sheet, so a column from another sheet is unnameable.
export type EndpointsAll = {
  [FN in ColumnFullName]?: Endpoint<SheetNameOf<FN>>;
};

// The framework owns the config sheets' entries, so the app's map can't name one.
export type Endpoints = {
  [
    FN in ColumnFullName as SheetNameOf<FN> extends FloorTabName ? never : FN
  ]?: Endpoint<SheetNameOf<FN>>;
};
