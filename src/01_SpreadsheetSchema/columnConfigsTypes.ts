import type {
  CodebaseNameDelimiter,
  NotEmpty,
} from "../00_Source/CellValues/cellValues";
import { lazy } from "../utils/lazy";
import { type FlattenTwoLevels, type KeyedMap, Obj } from "../utils/Obj";
import { Val } from "../utils/Val";
import { type Configs, installedConfigs } from "./configRegister";
import type { ColumnConfigsGeneric, ColumnConfigStored } from "./makeConfigs";
import {
  configSheetNames,
  getSheetTraitByName,
  type SheetNameSimple,
} from "./sheetConfigsTypes";
import { type Value, type ValueName, type ValueSchema } from "./valueSchemas";

export type ColumnConfigs = Configs["columnConfigs"];

// The generated literal read by an arbitrary name, where an entry may be absent.
export function columnConfigsByName(): ColumnConfigsGeneric {
  return columnConfigs();
}

function columnConfigs(): ColumnConfigs {
  return installedConfigs().columnConfigs;
}

export type ColumnName<SN extends SheetNameSimple = SheetNameSimple> =
  SN extends SheetNameSimple ? keyof ColumnConfigs[SN] : never;

// Distributes over SN; indexing a union of sheets by a union of column names collapses to never.
export type ColumnValueName<
  SN extends SheetNameSimple,
  CN extends ColumnName<SN>,
> = SN extends SheetNameSimple
  ? CN extends keyof ColumnConfigs[SN]
    ? ColumnConfigs[SN][CN]["valueName" & keyof ColumnConfigs[SN][CN]]
    : never
  : never;

export type ColumnIsFormula<
  SN extends SheetNameSimple,
  CN extends ColumnName<SN>,
> = SN extends SheetNameSimple
  ? CN extends keyof ColumnConfigs[SN]
    ? ColumnConfigs[SN][CN]["isFormula" & keyof ColumnConfigs[SN][CN]]
    : never
  : never;

export type ColumnEmptyValueAllowed<
  SN extends SheetNameSimple,
  CN extends ColumnName<SN>,
> = SN extends SheetNameSimple
  ? CN extends keyof ColumnConfigs[SN]
    ? ColumnConfigs[SN][CN]["emptyValueAllowed" & keyof ColumnConfigs[SN][CN]]
    : never
  : never;

export type ColumnNameFiltered<
  SN extends SheetNameSimple,
  VN extends ValueName = ValueName,
  IF extends boolean = boolean,
> = SN extends SheetNameSimple
  ? {
      [CN in ColumnName<SN>]: ColumnValueName<SN, CN> extends VN
        ? ColumnIsFormula<SN, CN> extends IF
          ? CN
          : never
        : never;
    }[ColumnName<SN>] &
      ColumnName<SN>
  : never;

export interface ColumnConfig<
  VN extends ValueName = ValueName,
> extends ColumnConfigStored<VN> {
  columnName: string;
}
export type ColumnConfigAt<
  SN extends SheetNameSimple,
  CN extends ColumnName<SN>,
> = ColumnConfig<ColumnValueName<SN, CN>>;

export type ColumnValueSchema<
  SN extends SheetNameSimple,
  CN extends ColumnName<SN>,
> = ValueSchema<ColumnValueName<SN, CN>>;

export type ColumnValue<
  SN extends SheetNameSimple,
  CN extends ColumnName<SN>,
> = Value<ColumnValueName<SN, CN>>;

// What the column's own Empty value allowed box declares the unmarked read to mean.
export type ColumnValueDeclared<
  SN extends SheetNameSimple,
  CN extends ColumnName<SN>,
> =
  ColumnEmptyValueAllowed<SN, CN> extends true
    ? ColumnValue<SN, CN>
    : NotEmpty<ColumnValue<SN, CN>>;

export type SheetDataValues<
  SN extends SheetNameSimple,
  CS extends ColumnName<SN> = ColumnName<SN>,
> = {
  [CN in CS]: ColumnValue<SN, CN>;
};

// Every writable column but the generated id: the bag a complete append must fill.
export type SheetDataValuesAll<SN extends SheetNameSimple> = SheetDataValues<
  SN,
  Exclude<ColumnNameFiltered<SN, ValueName, false>, "id">
>;

export function getSheetColumnNames<SN extends SheetNameSimple>(
  sheetName: SN,
): ColumnName<SN>[] {
  return Obj.keys(columnConfigs()[sheetName]) as unknown as ColumnName<SN>[];
}

// columnConfig isn't actually very unique. The only unique
export function getColumnTraitByName<
  TN extends SheetNameSimple,
  CN extends ColumnName<TN>,
  TK extends keyof ColumnConfigAt<TN, CN>,
>(sheetName: TN, columnName: CN, key: TK): ColumnConfigAt<TN, CN>[TK] {
  if (key === "columnName") {
    return columnName as ColumnConfigAt<TN, CN>[TK];
  }
  return (columnConfigs()[sheetName][columnName] as ColumnConfigAt<TN, CN>)[
    key
  ];
}

export type TableColumnConfigsById = KeyedMap<
  Record<string, ColumnConfigStored>,
  "columnId",
  "columnName"
>;

type ColumnConfigsByGidAndColId = Map<number, TableColumnConfigsById>;
function makeColumnConfigsByGidAndColId(): ColumnConfigsByGidAndColId {
  return configSheetNames().reduce((attrs, sheetName) => {
    const sheetGid = getSheetTraitByName(sheetName, "sheetGid");
    attrs.set(
      sheetGid,
      Obj.toKeyedMap(columnConfigs()[sheetName], "columnId", "columnName"),
    );
    return attrs;
  }, new Map() as ColumnConfigsByGidAndColId);
}

const columnConfigsByGidAndColId = lazy(makeColumnConfigsByGidAndColId);

export function getColumnTraitById<TK extends keyof ColumnConfig>(
  sheetGid: number,
  columnId: string,
  key: TK,
): ColumnConfig[TK] {
  const colTraits = Val.assert(
    columnConfigsByGidAndColId().get(sheetGid)?.get(columnId),
    `column attributes for sheetGid=${sheetGid}, columnId=${columnId}`,
  );
  return colTraits[key];
}
export function getSheetColumnIds(sheetGid: number): MapIterator<string> {
  return Val.assert(
    columnConfigsByGidAndColId().get(sheetGid),
    `column attributes for sheetGid=${sheetGid}`,
  ).keys();
}

export type MakeColumnFullName<
  SN extends SheetNameSimple,
  CN extends ColumnName<SN>,
> = `${SN}${CodebaseNameDelimiter}${CN & string}`;

type ColumnConfigsFlat = FlattenTwoLevels<
  ColumnConfigs,
  CodebaseNameDelimiter,
  "sheetName",
  "columnName"
>;
type ColumnFullNameAll = keyof ColumnConfigsFlat & string;

// Absolute addressing: one correlated key, so a filtered subset of columns is expressible.
export type ColumnFullName<
  VN extends ValueName = ValueName,
  IF extends boolean = boolean,
> = {
  [FN in ColumnFullNameAll]: ColumnConfigsFlat[FN]["valueName"] extends VN
    ? ColumnConfigsFlat[FN]["isFormula"] extends IF
      ? FN
      : never
    : never;
}[ColumnFullNameAll];

export type SheetNameOf<FN extends ColumnFullName> =
  ColumnConfigsFlat[FN]["sheetName"];
export type ColumnNameOf<FN extends ColumnFullName> =
  ColumnConfigsFlat[FN]["columnName"] & ColumnName<SheetNameOf<FN>>;
export type ValueNameOf<FN extends ColumnFullName> =
  ColumnConfigsFlat[FN]["valueName"];
export type ValueOf<FN extends ColumnFullName> = Value<ValueNameOf<FN>>;
