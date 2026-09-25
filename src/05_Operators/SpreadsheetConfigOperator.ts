import { installedConfigs } from "../01_SpreadsheetSchema/configRegister";
import { makeImportLine } from "../01_SpreadsheetSchema/makeConfigs";
import {
  spreadsheetConfigColumnLabel,
  spreadsheetConfigIndexHeaders,
  spreadsheetConfigTextHeaders,
} from "../01_SpreadsheetSchema/spreadsheetConfigFields";
import type { LiveSpreadsheetConfig } from "../01_SpreadsheetSchema/spreadsheetConfigTypes";
import { Obj } from "../utils/Obj";
import { Str } from "../utils/Str";
import { spreadsheetConfigFileSource } from "./configFileSource";
import { GenericSheetOperator } from "./GenericSheetOperator";
import {
  type ConfigSyncState,
  type OperatorProps,
  SpreadsheetBaseOperator,
} from "./SpreadsheetBaseOperator";

const fieldColumnNames = [
  ...Obj.values(spreadsheetConfigTextHeaders),
  ...Obj.values(spreadsheetConfigIndexHeaders),
].map((header) => Str.sentenceToCamelCase(header));

const {
  idHeader: editableIdHeader,
  nameHeader: editableNameHeader,
  ...fixedTextHeaders
} = spreadsheetConfigTextHeaders;

export class SpreadsheetConfigOperator extends GenericSheetOperator<"spreadsheetConfig"> {
  constructor(props: OperatorProps) {
    super({
      sheetName: "spreadsheetConfig",
      ...props,
    });
  }
  static init(): SpreadsheetConfigOperator {
    return new SpreadsheetConfigOperator(
      SpreadsheetBaseOperator.initOperatorProps(),
    );
  }
  get spreadsheetConfigSync(): ConfigSyncState["spreadsheetConfigSync"] {
    return this.configSyncState.spreadsheetConfigSync;
  }
  fetchLiveConfig(): LiveSpreadsheetConfig {
    this.sheet.prepFetchColumnsSpecific(
      [this.schema.topDataRowIdx],
      ...fieldColumnNames,
    );
    this.ss.fetchAllPrepped({ skipFetchingProperties: true });
    const firstDataRow = this.sheet.topRow;
    const textValues = Obj.mapValues(spreadsheetConfigTextHeaders, (header) =>
      firstDataRow.value(Str.sentenceToCamelCase(header)),
    );
    const indexValuesBase1 = Obj.mapValues(
      spreadsheetConfigIndexHeaders,
      (header) => firstDataRow.value(Str.sentenceToCamelCase(header)),
    );
    validateFixedLayoutValuesUnchanged(textValues, indexValuesBase1);
    const liveConfig: LiveSpreadsheetConfig = {
      ...textValues,
      ...Obj.mapValues(indexValuesBase1, (value) => value - 1),
    };
    this.spreadsheetConfigSync.liveConfig = liveConfig;
    return liveConfig;
  }
  toFileSource(makeConfigsImport: string): string {
    const liveConfig = this.spreadsheetConfigSync.liveConfig;
    if (liveConfig === undefined) {
      throw new Error(
        "SpreadsheetConfigOperator has not yet fetched the live Spreadsheet Config.",
      );
    }
    return [
      `${makeImportLine("makeSpreadsheetConfig", makeConfigsImport)}`,
      ``,
      `export const spreadsheetConfig = makeSpreadsheetConfig(${spreadsheetConfigFileSource(
        liveConfig,
      )} as const);`,
      ``,
    ].join("\n");
  }
  validateExactlyOneDataRow(): void {
    const dataRowCount = this.sheet.raw.dataRowCountAfterFlush;
    if (dataRowCount !== 1) {
      throw new Error(
        `Spreadsheet Config Table must have exactly one data row; found ${dataRowCount}.`,
      );
    }
  }
}

// Compared in base 1, so a refusal shows each cell as it reads on the sheet.
function validateFixedLayoutValuesUnchanged(
  textValues: Record<keyof typeof spreadsheetConfigTextHeaders, string>,
  indexValuesBase1: Record<keyof typeof spreadsheetConfigIndexHeaders, number>,
): void {
  const { spreadsheetConfig } = installedConfigs();
  const changed = [
    ...changedLayoutValueLines(
      fixedTextHeaders,
      textValues,
      (key) => spreadsheetConfig[key],
    ),
    ...changedLayoutValueLines(
      spreadsheetConfigIndexHeaders,
      indexValuesBase1,
      (key) => spreadsheetConfig[key] + 1,
    ),
  ];
  if (changed.length > 0) {
    throw new Error(
      `Spreadsheet Config layout values other than "${editableIdHeader}" and "${editableNameHeader}" are fixed; put back:\n${changed.join("\n")}`,
    );
  }
}

function changedLayoutValueLines<LK extends string>(
  headers: Record<LK, string>,
  liveValues: NoInfer<Record<LK, string | number>>,
  expectedValue: (key: LK) => string | number,
): string[] {
  return Obj.keys(headers).flatMap((key) => {
    const live = liveValues[key];
    const expected = expectedValue(key);
    return live === expected
      ? []
      : [
          `${spreadsheetConfigColumnLabel(headers[key])} is ${JSON.stringify(live)}; expected ${JSON.stringify(expected)}.`,
        ];
  });
}
