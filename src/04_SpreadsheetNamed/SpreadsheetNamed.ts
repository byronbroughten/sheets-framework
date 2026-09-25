import type { SheetName } from "../01_SpreadsheetSchema/sheetConfigsTypes.js";
import { SpreadsheetSchema } from "../01_SpreadsheetSchema/SpreadsheetSchema";
import type { FindReplaceProps } from "../02_SpreadsheetRaw/ClassTypes/StateRaw";
import { SpreadsheetRaw } from "../02_SpreadsheetRaw/SpreadsheetRaw.js";
import type { SheetIdentified } from "../03_SpreadsheetIdentified/SheetIdentified";
import type { GatherDataPrerequisitesProps } from "../03_SpreadsheetIdentified/SheetMetaIdentified";
import { SpreadsheetIdentified } from "../03_SpreadsheetIdentified/SpreadsheetIdentified.js";
import { Obj } from "../utils/Obj.js";
import { SerialDate } from "../utils/SerialDate.js";
import { Tim } from "../utils/Tim.js";
import { Val } from "../utils/Val.js";
import { SpreadsheetBaseNamed } from "./ClassBases/SpreadsheetBaseNamed.js";
import { SheetMetaNamed } from "./SheetMetaNamed.js";
import { SheetNamed } from "./SheetNamed.js";
import type { SheetNameByGroup } from "./SheetNameGroups.js";
import {
  type ColumnSpecifierNamed,
  type FetchColumnSpecifierNamed,
  type FetchPropsNamed,
  type FetchPropsStandardNamed,
  type NamedSheets,
  type RowSpecifierName,
  type SheetColumnNamesStandard,
} from "./Types/NamedState.js";

export class SpreadsheetNamed extends SpreadsheetBaseNamed {
  static init(): SpreadsheetNamed {
    return new SpreadsheetNamed(SpreadsheetNamed.initSpreadsheetNamedProps());
  }
  get schema(): SpreadsheetSchema {
    return new SpreadsheetSchema();
  }
  get raw(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
  get identified(): SpreadsheetIdentified {
    return new SpreadsheetIdentified(this.spreadsheetIdentifiedProps);
  }
  today(): SerialDate {
    return SerialDate.fromInstant(new Date(), this.raw.timeZone);
  }
  now(): string {
    return Tim.nowTimestamp(this.raw.timeZone);
  }
  get serialDate(): typeof SerialDate & { today(): SerialDate } {
    return { ...SerialDate, today: () => this.today() };
  }
  sheet<TN extends SheetName>(sheetName: TN): SheetNamed<TN> {
    return new SheetNamed({
      sheetName,
      ...this.spreadsheetNamedProps,
    });
  }
  sheetMeta<TN extends SheetName>(sheetName: TN): SheetMetaNamed<TN> {
    return new SheetMetaNamed({
      sheetName,
      ...this.spreadsheetNamedProps,
    });
  }
  sheets<TN extends SheetName>(...sheetNames: TN[]): NamedSheets<TN> {
    return sheetNames.reduce((acc, sheetName) => {
      acc[sheetName] = this.sheet(sheetName);
      return acc;
    }, {} as NamedSheets<TN>);
  }
  get activeSheetNames(): SheetName[] {
    return this.identified.activeSheets.map((sheet) => sheet.sheetName);
  }
  get activeSheets(): SheetNamed<SheetName>[] {
    return this.activeSheetNames.map((sheetName) => this.sheet(sheetName));
  }
  fetchAllPrepped(props: GatherDataPrerequisitesProps = {}): SpreadsheetNamed {
    this.identified.fetchAllPrepped(props);
    return this;
  }
  fetch<SN extends SheetName>(
    ...props: FetchPropsNamed<SN>[]
  ): NamedSheets<SN> {
    const standardizedProps = this._standardizeProps(props);
    this._prepFetchStandardizedProps(standardizedProps);
    this.fetchAllPrepped();
    const sheetNames = sheetNamesFromReqProps(standardizedProps);
    return this.sheets(...sheetNames);
  }
  private _standardizeProps<SN extends SheetName>(
    propsArr: FetchPropsNamed<SN>[],
  ): FetchPropsStandardNamed<SN>[] {
    return propsArr.map((props) => {
      const { rowSpecifier } = props;
      const columnSpecifiers = this._standardizeColumnSpecifiers(props);
      return {
        rowSpecifier,
        sheetColumnNames: columnSpecifiers,
      };
    });
  }

  private _standardizeColumnSpecifiers<SN extends SheetName>(
    columnSpecifier: FetchColumnSpecifierNamed<SN>,
  ): SheetColumnNamesStandard<SN> {
    if (columnSpecifier.sheetColumnMode === "all") {
      return this._allColumnNamesOf(
        this.schema.sheetNames,
      ) as SheetColumnNamesStandard<SN>;
    } else if (columnSpecifier.sheetColumnMode === "allColumns") {
      const sheetNames = Array.isArray(columnSpecifier.sheetNames)
        ? columnSpecifier.sheetNames
        : [columnSpecifier.sheetNames];
      return this._allColumnNamesOf(sheetNames);
    } else if (columnSpecifier.sheetColumnMode === "specific") {
      const sheetColumnNames = columnSpecifier.sheetColumnNames;
      return Obj.keys(sheetColumnNames).reduce((acc, sheetName) => {
        const schema = this.schema.sheetByName(sheetName);
        acc[sheetName] = schema.columnSpecifierToStandard(
          Val.assert<ColumnSpecifierNamed<SN>>(
            sheetColumnNames[sheetName],
            `sheetColumnNames[${sheetName}]`,
          ),
        );
        return acc;
      }, {} as SheetColumnNamesStandard<SN>);
    } else {
      throw new Error(
        `Invalid sheetColumnMode: ${
          (columnSpecifier as FetchColumnSpecifierNamed<SN>).sheetColumnMode
        }. Must be a valid ColumnMode.`,
      );
    }
  }

  private _allColumnNamesOf<SN extends SheetName>(
    sheetNames: readonly SN[],
  ): SheetColumnNamesStandard<SN> {
    return sheetNames.reduce((acc, sheetName) => {
      acc[sheetName] = this.schema.sheetByName(sheetName).columnNames;
      return acc;
    }, {} as SheetColumnNamesStandard<SN>);
  }
  private _prepFetchStandardizedProps(
    propsArr: FetchPropsStandardNamed<SheetName>[],
  ): void {
    propsArr.forEach((props) => this._prepFetchStandardProps(props));
  }
  private _prepFetchStandardProps({
    rowSpecifier,
    sheetColumnNames,
  }: FetchPropsStandardNamed): void {
    const specifiers =
      typeof rowSpecifier === "string" ? [rowSpecifier] : rowSpecifier;
    Obj.keys(sheetColumnNames).forEach((sheetName) => {
      const columnNames = Val.assert(
        sheetColumnNames[sheetName],
        `sheetColumnNames[${sheetName}]`,
      );
      const namedSheet = this.sheet(sheetName);
      const identifiedSheet = namedSheet.identified;
      columnNames.forEach((columnName) => {
        const columnId = namedSheet.schema.columnByName(columnName).columnId;
        specifiers.forEach((specifier) => {
          prepFetchRowSpecifier(identifiedSheet, specifier, columnId);
        });
      });
    });
  }
  get sheetsOfSchema(): SheetNamed<SheetName>[] {
    return this.schema.sheetNames.map((sheetName) => this.sheet(sheetName));
  }
  batchUpdateGSheets(): void {
    this.raw.batchUpdateGSheets();
  }
  // The scope is explicit here because allSheets has no narrower home.
  findReplace(props: FindReplaceProps): this {
    this.raw.findReplace(props);
    return this;
  }
  discardQueuedChanges(): this {
    this.raw.discardQueuedChanges();
    return this;
  }
  fetchAllSheetProperties(): this {
    this.raw.fetchAllSheetProperties();
    return this;
  }
  ensureAllSheetPropertiesAreFetched(): this {
    this.raw.ensureAllSheetPropertiesAreFetched();
    return this;
  }
  fillMissingRowIds(): void {
    // could potentially be reconfigured to not rely on the schema.
    const idSheets = this._sheetsWithRowIds();
    idSheets.forEach((sheet) => {
      sheet.column("id").prepFetchFull();
    });
    this.fetchAllPrepped();
    idSheets.forEach((sheet) => {
      sheet.column("id").emptyActiveCellsToDefualt();
    });
  }
  private _sheetsWithRowIds(): SheetNamed<SheetNameByGroup<"hasIdColumn">>[] {
    return this.sheetsOfSchema.filter((sheet) => {
      const hasIdCol = sheet.schema.trait("hasIdColumn");
      const idPrefix = sheet.schema.trait("idPrefix");
      if (hasIdCol) {
        if (!idPrefix) {
          throw new Error(
            `Cannot fill missing row IDs in sheet "${sheet}" because it does not have an "ID prefix" defined in the schema.`,
          );
        }
        return true;
      }
      return false;
    }) as SheetNamed<SheetNameByGroup<"hasIdColumn">>[];
  }
}

function sheetNamesFromReqProps<SN extends SheetName>(
  propsArr: FetchPropsStandardNamed<SN>[],
): Set<SN> {
  return propsArr.reduce((sheetNames, props) => {
    return sheetNames.add(...Obj.keys(props.sheetColumnNames));
  }, new Set() as Set<SN>);
}

function prepFetchRowSpecifier(
  sheet: SheetIdentified,
  rowSpecifier: RowSpecifierName,
  columnId: string,
): void {
  const schema = sheet.schema;
  const column = sheet.column(columnId);
  switch (rowSpecifier) {
    case "activeRows":
    case "data":
      column.prepFetchFull();
      break;
    case "topDatum":
      column.cell(schema.topDataRowIdx).prepFetch();
      break;
    case "actions":
      column.cell(schema.actionRowIndex).prepFetch();
      break;
    case "columnIds":
      column.cell(schema.colIdRowIndex).prepFetch();
      break;
    case "headers":
      column.cell(schema.tableHeaderRowIndex).prepFetch();
      break;
    case "all":
      (["headers", "actions", "columnIds", "data"] as const).forEach(
        (specifier) => prepFetchRowSpecifier(sheet, specifier, columnId),
      );
      break;
    default:
      throw new Error(
        `Invalid rowSpecifier: ${rowSpecifier as string}. Must be a valid RowSpecifierName.`,
      );
  }
}
