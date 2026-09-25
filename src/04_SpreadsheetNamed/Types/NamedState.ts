import type { ColumnName } from "../../01_SpreadsheetSchema/columnConfigsTypes";
import type { SheetName } from "../../01_SpreadsheetSchema/sheetConfigsTypes";
import type { StrictOmit } from "../../utils/Obj";
import type { SheetNamed } from "../SheetNamed";

type SheetColumnNames<SN extends SheetName> = {
  [S in SN]?: ColumnSpecifierNamed<SN>;
};

export interface FetchSpecifierObjNamed<SN extends SheetName = SheetName> {
  all: {
    rowSpecifier: RowSpecifier;
    sheetColumnMode: "all";
  };
  allColumns: {
    rowSpecifier: RowSpecifier;
    sheetColumnMode: "allColumns";
    sheetNames: SN | SN[];
  };
  specific: {
    rowSpecifier: RowSpecifier;
    sheetColumnMode: "specific";
    sheetColumnNames: SheetColumnNames<SN>;
  };
}

type FetchColumnsSpecifierObjNamed<SN extends SheetName = SheetName> = {
  [S in keyof FetchSpecifierObjNamed<SN>]: StrictOmit<
    FetchSpecifierObjNamed<SN>[S],
    "rowSpecifier"
  >;
};

export type FetchColumnSpecifierNamed<SN extends SheetName> =
  FetchColumnsSpecifierObjNamed<SN>[ColumnMode];

type ColumnMode = keyof FetchSpecifierObjNamed;

export type FetchPropsNamed<SN extends SheetName> =
  FetchSpecifierObjNamed<SN>[ColumnMode];

export type SheetColumnNamesStandard<SN extends SheetName> = {
  [S in SN]?: ColumnName<SN>[];
};

export interface FetchPropsStandardNamed<SN extends SheetName = SheetName> {
  rowSpecifier: RowSpecifier;
  sheetColumnNames: SheetColumnNamesStandard<SN>;
}

type RowSpecifier = RowSpecifierName | RowSpecifierName[];
export type RowSpecifierBySchemaName = Exclude<RowSpecifierName, "activeRows">;

export const rowSpecifierNames = [
  "all",
  "activeRows",
  "data",
  "topDatum",
  "actions",
  "columnIds",
  "headers",
] as const;
export type RowSpecifierName = (typeof rowSpecifierNames)[number];
export function isRowName(value: unknown): value is RowSpecifierName {
  return (
    typeof value === "string" &&
    rowSpecifierNames.includes(value as RowSpecifierName)
  );
}

export type ColumnSpecifierNamed<TN extends SheetName> =
  ColumnName<TN> | ColumnName<TN>[] | "allColumns";

export type NamedSheets<TN extends SheetName> = {
  [T in TN]: SheetNamed<T>;
};
