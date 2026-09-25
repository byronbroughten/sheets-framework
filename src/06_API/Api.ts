import type { SheetChange } from "../00_Source/PlatformEvents/sheetChange";
import type { SheetEdit } from "../00_Source/PlatformEvents/sheetEdit";
import type { ColumnSchema } from "../01_SpreadsheetSchema/ColumnSchema";
import {
  type Configs,
  installConfigs,
} from "../01_SpreadsheetSchema/configRegister";
import { ssConfigGet } from "../01_SpreadsheetSchema/spreadsheetConfigTypes";
import { SpreadsheetSchema } from "../01_SpreadsheetSchema/SpreadsheetSchema";
import { SpreadsheetIdentified } from "../03_SpreadsheetIdentified/SpreadsheetIdentified";
import {
  SpreadsheetBaseNamed,
  type SpreadsheetNamedProps,
} from "../04_SpreadsheetNamed/ClassBases/SpreadsheetBaseNamed";
import { ConfigSheetFloor } from "../05_Operators/ConfigSheetFloor";
import type { FloorNotice } from "../05_Operators/ConfigSheetFloor/floorChangeNotice";
import { EndpointRun } from "./EndpointRun";
import type { Endpoints, EndpointsAll } from "./Endpoints";
import { frameworkEndpoints } from "./frameworkEndpoints";

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
    this.endpoints = {
      ...endpoints,
      ...frameworkEndpoints,
    };
  }
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
    installSource();
    Api.init(endpoints).handleSheetEdit(edit);
  }
  static handleSheetChange(
    { configs }: AppSetup,
    change: SheetChange | undefined,
    installSource: () => void,
  ): FloorNotice | undefined {
    if (change === undefined) return undefined;
    installConfigs(configs);
    installSource();
    return ConfigSheetFloor.init().changeNotice(change);
  }
  get schema(): SpreadsheetSchema {
    return new SpreadsheetSchema();
  }
  get ssi(): SpreadsheetIdentified {
    return new SpreadsheetIdentified(this.spreadsheetIdentifiedProps);
  }
  static isSuspectedApiCall(edit: SheetEdit): boolean {
    return (
      (edit.value === "TRUE" || edit.value === "FALSE") &&
      edit.rowIndexBase0 === ssConfigGet("actionRowIndexBase0")
    );
  }
  handleSheetEdit({ sheetGid, colIndexBase0, value }: SheetEdit): void {
    if (!this.schema.isInSheetGids(sheetGid)) {
      return;
    }
    const sheet = this.ssi.sheetMeta(sheetGid).ensureColumnIdsAreFetched();
    if (!sheet.isTableColIndex(colIndexBase0)) {
      return;
    }
    const columnId = sheet.columnIdByIndex(colIndexBase0);
    if (columnId === "") {
      return;
    }
    this._runEndpoint(sheet.schema.columnById(columnId), value === "TRUE");
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
      sheetName: entryColumn.sheetName,
      entryColumnName: entryColumn.columnName,
      endpoint,
    }).run(isChecked);
  }
}
