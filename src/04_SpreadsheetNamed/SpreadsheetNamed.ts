import type { TableName } from "../01_SpreadsheetSchema/sheetConfigsTypes.js";
import { SpreadsheetSchema } from "../01_SpreadsheetSchema/SpreadsheetSchema";
import type { FindReplaceProps } from "../02_SpreadsheetRaw/ClassTypes/StateRaw";
import { SpreadsheetRaw } from "../02_SpreadsheetRaw/SpreadsheetRaw.js";
import type { GatherDataPrerequisitesProps } from "../03_SpreadsheetIdentified/SheetMetaIdentified";
import { SpreadsheetIdentified } from "../03_SpreadsheetIdentified/SpreadsheetIdentified.js";
import type { TableIdentified } from "../03_SpreadsheetIdentified/TableIdentified";
import { Obj } from "../utils/Obj";
import { SerialDate } from "../utils/SerialDate";
import { SerialDateTime } from "../utils/SerialDateTime";
import { Val } from "../utils/Val";
import { SpreadsheetBaseNamed } from "./ClassBases/SpreadsheetBaseNamed.js";
import { SheetMetaNamed } from "./SheetMetaNamed.js";
import { TableNamed } from "./TableNamed.js";
import type { SheetNameByGroup } from "./TableNameGroups.js";
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
    return SerialDateTime.nowTimestamp(this.raw.timeZone);
  }
  get serialDate(): typeof SerialDate & { today(): SerialDate } {
    return { ...SerialDate, today: () => this.today() };
  }
  sheetMeta<TN extends TableName>(tableName: TN): SheetMetaNamed<TN> {
    return new SheetMetaNamed({
      tableName,
      ...this.spreadsheetNamedProps,
    });
  }
  table<TN extends TableName>(tableName: TN): TableNamed<TN> {
    return new TableNamed({
      tableName,
      ...this.spreadsheetNamedProps,
    });
  }
  tables<TN extends TableName>(...tableNames: TN[]): NamedSheets<TN> {
    return tableNames.reduce((acc, tableName) => {
      acc[tableName] = this.table(tableName);
      return acc;
    }, {} as NamedSheets<TN>);
  }
  get activeSheetNames(): TableName[] {
    return this.identified.activeSheets.map((sheet) => sheet.tableName);
  }
  get activeSheets(): TableNamed<TableName>[] {
    return this.activeSheetNames.map((tableName) => this.table(tableName));
  }
  fetchAllPrepped(props: GatherDataPrerequisitesProps = {}): SpreadsheetNamed {
    this.identified.fetchAllPrepped(props);
    return this;
  }
  fetch<TN extends TableName>(
    ...props: FetchPropsNamed<TN>[]
  ): NamedSheets<TN> {
    const standardizedProps = this._standardizeProps(props);
    this._prepFetchStandardizedProps(standardizedProps);
    this.fetchAllPrepped();
    const sheetNames = sheetNamesFromReqProps(standardizedProps);
    return this.tables(...sheetNames);
  }
  private _standardizeProps<TN extends TableName>(
    propsArr: FetchPropsNamed<TN>[],
  ): FetchPropsStandardNamed<TN>[] {
    return propsArr.map((props) => {
      const { rowSpecifier } = props;
      const columnSpecifiers = this._standardizeColumnSpecifiers(props);
      return {
        rowSpecifier,
        sheetColumnNames: columnSpecifiers,
      };
    });
  }

  private _standardizeColumnSpecifiers<TN extends TableName>(
    columnSpecifier: FetchColumnSpecifierNamed<TN>,
  ): SheetColumnNamesStandard<TN> {
    if (columnSpecifier.sheetColumnMode === "all") {
      return this._allColumnNamesOf(
        this.schema.sheetNames,
      ) as SheetColumnNamesStandard<TN>;
    } else if (columnSpecifier.sheetColumnMode === "allColumns") {
      const sheetNames = Array.isArray(columnSpecifier.sheetNames)
        ? columnSpecifier.sheetNames
        : [columnSpecifier.sheetNames];
      return this._allColumnNamesOf(sheetNames);
    } else if (columnSpecifier.sheetColumnMode === "specific") {
      const sheetColumnNames = columnSpecifier.sheetColumnNames;
      return Obj.keys(sheetColumnNames).reduce((acc, tableName) => {
        const schema = this.schema.sheetByName(tableName);
        acc[tableName] = schema.columnSpecifierToStandard(
          Val.assert<ColumnSpecifierNamed<TN>>(
            sheetColumnNames[tableName],
            `sheetColumnNames[${tableName}]`,
          ),
        );
        return acc;
      }, {} as SheetColumnNamesStandard<TN>);
    } else {
      throw new Error(
        `Invalid sheetColumnMode: ${
          (columnSpecifier as FetchColumnSpecifierNamed<TN>).sheetColumnMode
        }. Must be a valid ColumnMode.`,
      );
    }
  }

  private _allColumnNamesOf<TN extends TableName>(
    sheetNames: readonly TN[],
  ): SheetColumnNamesStandard<TN> {
    return sheetNames.reduce((acc, tableName) => {
      acc[tableName] = this.schema.sheetByName(tableName).columnNames;
      return acc;
    }, {} as SheetColumnNamesStandard<TN>);
  }
  private _prepFetchStandardizedProps(
    propsArr: FetchPropsStandardNamed<TableName>[],
  ): void {
    propsArr.forEach((props) => this._prepFetchStandardProps(props));
  }
  private _prepFetchStandardProps({
    rowSpecifier,
    sheetColumnNames,
  }: FetchPropsStandardNamed): void {
    const specifiers =
      typeof rowSpecifier === "string" ? [rowSpecifier] : rowSpecifier;
    Obj.keys(sheetColumnNames).forEach((tableName) => {
      const columnNames = Val.assert(
        sheetColumnNames[tableName],
        `sheetColumnNames[${tableName}]`,
      );
      const namedSheet = this.table(tableName);
      const identifiedSheet = namedSheet.identified;
      columnNames.forEach((columnName) => {
        const columnId = namedSheet.schema.columnByName(columnName).columnId;
        specifiers.forEach((specifier) => {
          prepFetchRowSpecifier(identifiedSheet, specifier, columnId);
        });
      });
    });
  }
  get sheetsOfSchema(): TableNamed<TableName>[] {
    return this.schema.sheetNames.map((tableName) => this.table(tableName));
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
  private _sheetsWithRowIds(): TableNamed<SheetNameByGroup<"hasIdColumn">>[] {
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
    }) as TableNamed<SheetNameByGroup<"hasIdColumn">>[];
  }
}

function sheetNamesFromReqProps<TN extends TableName>(
  propsArr: FetchPropsStandardNamed<TN>[],
): Set<TN> {
  return new Set(propsArr.flatMap((props) => Obj.keys(props.sheetColumnNames)));
}

function prepFetchRowSpecifier(
  sheet: TableIdentified,
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
      column.cell(0).prepFetch();
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
