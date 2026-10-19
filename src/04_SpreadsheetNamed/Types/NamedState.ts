import type { StrictOmit } from "@byronbroughten/utils/obj";

import type { ColumnName } from "../../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import type { TableName } from "../../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import type { TableNamed } from "../TableNamed";

type SheetColumnNames<TN extends TableName> = {
  [S in TN]?: ColumnSpecifierNamed<TN>;
};

export interface FetchSpecifierObjNamed<TN extends TableName = TableName> {
  all: {
    rowSpecifier: RowSpecifier;
    sheetColumnMode: "all";
  };
  allColumns: {
    rowSpecifier: RowSpecifier;
    sheetColumnMode: "allColumns";
    sheetNames: TN | TN[];
  };
  specific: {
    rowSpecifier: RowSpecifier;
    sheetColumnMode: "specific";
    sheetColumnNames: SheetColumnNames<TN>;
  };
}

type FetchColumnsSpecifierObjNamed<TN extends TableName = TableName> = {
  [S in keyof FetchSpecifierObjNamed<TN>]: StrictOmit<
    FetchSpecifierObjNamed<TN>[S],
    "rowSpecifier"
  >;
};

export type FetchColumnSpecifierNamed<TN extends TableName> =
  FetchColumnsSpecifierObjNamed<TN>[ColumnMode];

type ColumnMode = keyof FetchSpecifierObjNamed;

export type FetchPropsNamed<TN extends TableName> =
  FetchSpecifierObjNamed<TN>[ColumnMode];

export type SheetColumnNamesStandard<TN extends TableName> = {
  [S in TN]?: ColumnName<TN>[];
};

export interface FetchPropsStandardNamed<TN extends TableName = TableName> {
  rowSpecifier: RowSpecifier;
  sheetColumnNames: SheetColumnNamesStandard<TN>;
}

type RowSpecifier = RowSpecifierName | RowSpecifierName[];

export const rowSpecifierNames = [
  "all",
  "data",
  "topDatum",
  "actions",
  "columnIds",
  "headers",
] as const;
export type RowSpecifierName = (typeof rowSpecifierNames)[number];

export type ColumnSpecifierNamed<TN extends TableName> =
  ColumnName<TN> | ColumnName<TN>[] | "allColumns";

export type NamedSheets<TN extends TableName> = {
  [T in TN]: TableNamed<T>;
};
