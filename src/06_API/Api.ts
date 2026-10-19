import type { SheetEdit } from "../00_Source/PlatformEvents/sheetEdit";
import type { ColumnSchema } from "../01_SpreadsheetSchema/configReaders/ColumnSchema";
import {
  type Configs,
  installConfigs,
} from "../01_SpreadsheetSchema/configReaders/configRegister";
import { SpreadsheetSchema } from "../01_SpreadsheetSchema/configReaders/SpreadsheetSchema";
import { SpreadsheetIdentified } from "../03_SpreadsheetIdentified/SpreadsheetIdentified";
import {
  SpreadsheetBaseNamed,
  type SpreadsheetNamedProps,
} from "../04_SpreadsheetNamed/ClassBases/SpreadsheetBaseNamed";
import { EndpointRun } from "./EndpointRun";
import type { Endpoints, EndpointsAll } from "./Endpoints";
import { installEndpoints } from "./feedbackColumnIds";
import { withFrameworkEndpoints } from "./frameworkEndpoints";

interface ApiProps extends SpreadsheetNamedProps {
  endpoints: Endpoints;
}

export interface AppSetup {
  configs: Configs;
  endpoints: Endpoints;
}

export class Api extends SpreadsheetBaseNamed {
  readonly endpoints: EndpointsAll;
  constructor({ endpoints, ...rest }: ApiProps) {
    super(rest);
    this.endpoints = withFrameworkEndpoints(endpoints);
  }
  // Takes the installed feedback columns; the entry calls install them first.
  static init(endpoints: Endpoints): Api {
    return new Api({
      endpoints,
      ...SpreadsheetBaseNamed.initSpreadsheetNamedProps(),
    });
  }
  static handleSheetEdit(
    { configs, endpoints }: AppSetup,
    edit: SheetEdit,
    installSource: () => void,
  ): void {
    installConfigs(configs);
    if (!Api.isSuspectedApiCall(edit)) return;
    installEndpoints(endpoints);
    installSource();
    Api.init(endpoints).handleSheetEdit(edit);
  }
  get identified(): SpreadsheetIdentified {
    return new SpreadsheetIdentified(this.spreadsheetIdentifiedProps);
  }
  static isSuspectedApiCall(edit: SheetEdit): boolean {
    return (
      (edit.value === "TRUE" || edit.value === "FALSE") &&
      new SpreadsheetSchema().mayHoldActionCell(edit)
    );
  }
  handleSheetEdit(edit: SheetEdit): void {
    const table = this.identified.fetchTableWithActionCell(edit);
    if (table === undefined) {
      return;
    }
    const { colIndexBase0, value } = edit;
    const colIndex = table.raw.tableOrigin().colIndex(colIndexBase0);
    const columnId = table.columnResolver.columnIdAt(colIndex);
    if (columnId === "") {
      return;
    }
    this._runEndpoint(table.schema.columnById(columnId), value === "TRUE");
  }
  // An entry that doesn't run on uncheck is a button, so only ticking fires it.
  private _runEndpoint(entryColumn: ColumnSchema, isChecked: boolean): void {
    const endpoint = this.endpoints[entryColumn.fullName];
    if (!endpoint) {
      return;
    }
    if (!isChecked && !endpoint.runOnUncheck) {
      return;
    }
    // The full name is only known at runtime, so the run widens to every sheet.
    new EndpointRun({
      ...this.spreadsheetNamedProps,
      tableName: entryColumn.tableName,
      entryColumnName: entryColumn.columnName,
      endpoint,
    }).run(isChecked);
  }
}
