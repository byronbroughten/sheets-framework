import type { RgbColor } from "../00_Source/RawSource/RgbColor";
import type { ColumnName } from "../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import type { TableNameSimple } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import type { CellChange } from "../03_SpreadsheetIdentified/ClassTypes/StateIdentified";
import {
  TableBaseNamed,
  type TableNamedProps,
} from "../04_SpreadsheetNamed/ClassBases/TableBaseNamed";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import type { TableNamed } from "../04_SpreadsheetNamed/TableNamed";
import {
  type CheckboxColumnName,
  CheckboxColumnOperator,
} from "../05_Operators/CheckboxColumnOperator";
import type {
  ActionReturn,
  EndpointDispatched,
  FeedbackColumnName,
  RowReports,
  RunReport,
  RunStateReported,
} from "./Endpoints";

interface RunState {
  message: string;
  backgroundColor: RgbColor;
}

// Paired here so no path can show one state's colour beside another's message.
const runStates = {
  running: {
    message: "Running…",
    backgroundColor: { red: 1, green: 0.949, blue: 0.8 },
  },
  success: {
    message: "Succeeded",
    backgroundColor: { red: 0.851, green: 0.918, blue: 0.827 },
  },
  // Between running's yellow and failure's red, so the states read as a scale.
  warning: {
    message: "Warning",
    backgroundColor: { red: 0.99, green: 0.85, blue: 0.7 },
  },
  failure: {
    message: "Failed",
    backgroundColor: { red: 0.957, green: 0.8, blue: 0.8 },
  },
} as const satisfies Record<RunStateName, RunState>;

type RunStateName = RunStateReported | "running";

interface RunStateProps {
  startTime?: string;
  message?: string;
}

export interface EndpointRunProps<
  TN extends TableNameSimple,
> extends TableNamedProps<TN> {
  entryColumnName: ColumnName<TN>;
  endpoint: EndpointDispatched<TN>;
}

/**
 * Owns everything after Api dispatches: selector prune, running-state
 * flush, action try/catch, selection consume, outcome flush.
 * Api only decodes the edit event and looks up the endpoint map.
 * Action bodies live in businessEndpoints/<name>.ts; the registry is
 * businessEndpoints.ts. Run-report shape and flags: GLOSSARY.md.
 * docs/architecture/endpoint-dispatch.md
 */
export class EndpointRun<
  TN extends TableNameSimple = TableNameSimple,
> extends TableBaseNamed<TN> {
  readonly entryColumnName: ColumnName<TN>;
  readonly endpoint: EndpointDispatched<TN>;
  constructor({ entryColumnName, endpoint, ...props }: EndpointRunProps<TN>) {
    super(props);
    this.entryColumnName = entryColumnName;
    this.endpoint = endpoint;
  }
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  get table(): TableNamed<TN> {
    return this.ss.table(this.tableName);
  }
  run(isChecked: boolean): void {
    this._prepSelectorFetch();
    this.ss.fetchAllPrepped();
    const selectedRowIndexes = this._selectedRowIndexes();
    this._resetEntryCheckbox();
    if (this.endpoint.selector && selectedRowIndexes.length === 0) {
      this.ss.batchUpdateGSheets();
      Logger.log("No rows are selected, so the endpoint did not run.");
      return;
    }
    this._pruneToSelection(selectedRowIndexes);
    this._onRunSetup();
    try {
      this._validateOneRowSelected(selectedRowIndexes);
      const report = this.endpoint.action(this.ss, {
        selectedRowIndexes,
        isChecked,
      });
      this._clearSelection();
      this._applyActionReport(report);
    } catch (error) {
      this._onRunError(error);
    } finally {
      this.ss.batchUpdateGSheets();
    }
  }
  private _checkboxColumn(
    columnName: CheckboxColumnName<TN>,
  ): CheckboxColumnOperator<TN, CheckboxColumnName<TN>> {
    return new CheckboxColumnOperator({
      ...this.sheetNamedProps,
      columnName,
    });
  }
  // Prepped rather than fetched, so the selection rides the cycle already running.
  private _prepSelectorFetch(): void {
    const { selector } = this.endpoint;
    if (!selector) return;
    this._checkboxColumn(selector.column).column.prepFetchFull();
  }
  // No selector means every data row that holds data; a blank row is no record.
  private _selectedRowIndexes(): number[] {
    const { selector } = this.endpoint;
    if (!selector) return this.table.rowIndexesFullWithData;
    return this._checkboxColumn(selector.column).rowIndexesChecked;
  }
  // The entry cell is a button unless the endpoint also runs on unticking.
  private _resetEntryCheckbox(): void {
    if (this.endpoint.runOnUncheck) return;
    this.table
      .column(this.entryColumnName)
      .headCell("action")
      .updateValue(false);
  }
  // Unselected rows leave the working view, so every later read of working rows is the selection.
  private _pruneToSelection(selectedRowIndexes: number[]): void {
    if (!this.endpoint.selector) return;
    this.table.raw.removeRowsExcept(...selectedRowIndexes);
  }
  // The flush is what puts the running state on the sheet before the work runs.
  private _onRunSetup(): void {
    this._applyRunState("running", { startTime: this.ss.now() });
    this.ss.batchUpdateGSheets();
  }
  // First inside the `try`, so the refusal reports like any other failure and keeps the ticks.
  private _validateOneRowSelected(selectedRowIndexes: number[]): void {
    if (!this.endpoint.selector?.requireOneRow) return;
    if (selectedRowIndexes.length <= 1) return;
    throw new Error(
      // The sheet's own title, not its config name: the operator reads this cell.
      `This endpoint runs on one row of "${this.table.raw.sheet.title}" at a time, but ${selectedRowIndexes.length} are selected.`,
    );
  }
  // Inside the run's `try`, so an action that throws has its clearing discarded too.
  private _clearSelection(): void {
    const { selector } = this.endpoint;
    if (!selector || selector.retainSelection) return;
    this._checkboxColumn(selector.column).uncheckWorkingCells();
  }
  // A string is a success with that message, and nothing at all is a bare success.
  private _applyActionReport(report: ActionReturn): void {
    if (typeof report === "string") {
      this._applyRunState("success", { message: report });
      return;
    }
    const { runState, message, rows }: RunReport & { rows?: RowReports } =
      report ?? {};
    this._applyRunState(runState ?? "success", { message });
    rows?.forEach((rowReport, rowIndex) => {
      this._applyRowReport(rowIndex, rowReport);
    });
  }
  // The timestamp is written once at setup; a state change only recolours it.
  private _applyRunState(
    stateName: RunStateName,
    { startTime, message }: RunStateProps = {},
  ): void {
    const state = runStates[stateName];
    const { timeLastRan, runStatus } = this.endpoint;
    this._fillFeedbackColumn(timeLastRan, {
      value: startTime,
      backgroundColor: state.backgroundColor,
    });
    this._fillFeedbackColumn(runStatus, {
      value: message ?? state.message,
      backgroundColor: state.backgroundColor,
    });
  }
  private _fillFeedbackColumn(
    columnName: FeedbackColumnName<TN> | undefined,
    change: CellChange<"string">,
  ): void {
    if (!columnName) return;
    // Re-deriving the value type here would compose two mapped filters, at ~43k instantiations.
    const column = this.table.columnIdentified(columnName);
    if (this.endpoint.selector) {
      column.updateWorkingCells(change);
    } else {
      column.updateAllCells(change);
    }
  }
  // Per-cell and queued after the run-level fill, so it lands on top of it.
  private _applyRowReport(rowIndex: number, report: RunReport): void {
    this._validateIsDataRow(rowIndex);
    const state = runStates[report.runState ?? "success"];
    const { timeLastRan, runStatus } = this.endpoint;
    this._updateFeedbackCell(timeLastRan, rowIndex, {
      backgroundColor: state.backgroundColor,
    });
    this._updateFeedbackCell(runStatus, rowIndex, {
      value: report.message ?? state.message,
      backgroundColor: state.backgroundColor,
    });
  }
  // A key outside the data rows is a reporting bug, and would write somewhere surprising.
  private _validateIsDataRow(rowIndex: number): void {
    if (this.table.raw.rowIndexesFull.includes(rowIndex)) return;
    throw new Error(
      `This run cannot report into ${this.table.raw.rowLabel(rowIndex)}: it is not a data row of "${this.table.raw.sheet.title}".`,
    );
  }
  private _updateFeedbackCell(
    columnName: FeedbackColumnName<TN> | undefined,
    rowIndex: number,
    change: CellChange<"string">,
  ): void {
    if (!columnName) return;
    this.table.columnIdentified(columnName).cell(rowIndex).update(change);
  }
  // Queued changes are shared by reference, so a half-finished run must be dropped before status is written.
  private _onRunError(error: unknown): void {
    this.ss.discardQueuedChanges();
    this._applyRunState("failure", { message: String(error) });
    Logger.log(`Endpoint run failed: ${String(error)}`);
  }
}
