import type {
  CodebaseNameDelimiter,
  NotEmpty,
} from "../00_Source/CellValues/cellValues";
import { lazy } from "../utils/lazy";
import { type FlattenTwoLevels, type KeyedMap, Obj } from "../utils/Obj";
import { Val } from "../utils/Val";
import { type Configs, installedConfigs } from "./configRegister";
import type { ColumnConfigsGeneric, ColumnConfigStored } from "./makeConfigs";
import { configTableNames, type TableNameSimple } from "./tableConfigsTypes";
import { type Value, type ValueName, type ValueSchema } from "./valueSchemas";

export type ColumnConfigs = Configs["columnConfigs"];

// The generated literal read by an arbitrary name, where an entry may be absent.
export function columnConfigsByName(): ColumnConfigsGeneric {
  return columnConfigs();
}

function columnConfigs(): ColumnConfigs {
  return installedConfigs().columnConfigs;
}

export type ColumnName<TN extends TableNameSimple = TableNameSimple> =
  TN extends TableNameSimple ? keyof ColumnConfigs[TN] : never;

// Distributes over TN; indexing a union of sheets by a union of column names collapses to never.
export type ColumnValueName<
  TN extends TableNameSimple,
  CN extends ColumnName<TN>,
> = TN extends TableNameSimple
  ? CN extends keyof ColumnConfigs[TN]
    ? ColumnConfigs[TN][CN]["valueName" & keyof ColumnConfigs[TN][CN]]
    : never
  : never;

export type ColumnIsFormula<
  TN extends TableNameSimple,
  CN extends ColumnName<TN>,
> = TN extends TableNameSimple
  ? CN extends keyof ColumnConfigs[TN]
    ? ColumnConfigs[TN][CN]["isFormula" & keyof ColumnConfigs[TN][CN]]
    : never
  : never;

export type ColumnEmptyValueAllowed<
  TN extends TableNameSimple,
  CN extends ColumnName<TN>,
> = TN extends TableNameSimple
  ? CN extends keyof ColumnConfigs[TN]
    ? ColumnConfigs[TN][CN]["emptyValueAllowed" & keyof ColumnConfigs[TN][CN]]
    : never
  : never;

export type ColumnNameFiltered<
  TN extends TableNameSimple,
  VN extends ValueName = ValueName,
  IF extends boolean = boolean,
> = TN extends TableNameSimple
  ? {
      [CN in ColumnName<TN>]: ColumnValueName<TN, CN> extends VN
        ? ColumnIsFormula<TN, CN> extends IF
          ? CN
          : never
        : never;
    }[ColumnName<TN>] &
      ColumnName<TN>
  : never;

export interface ColumnConfig<
  VN extends ValueName = ValueName,
> extends ColumnConfigStored<VN> {
  columnName: string;
}
export type ColumnConfigAt<
  TN extends TableNameSimple,
  CN extends ColumnName<TN>,
> = ColumnConfig<ColumnValueName<TN, CN>>;

export type ColumnValueSchema<
  TN extends TableNameSimple,
  CN extends ColumnName<TN>,
> = ValueSchema<ColumnValueName<TN, CN>>;

export type ColumnValue<
  TN extends TableNameSimple,
  CN extends ColumnName<TN>,
> = Value<ColumnValueName<TN, CN>>;

// What the column's own Empty value allowed box declares the unmarked read to mean.
export type ColumnValueDeclared<
  TN extends TableNameSimple,
  CN extends ColumnName<TN>,
> =
  ColumnEmptyValueAllowed<TN, CN> extends true
    ? ColumnValue<TN, CN>
    : NotEmpty<ColumnValue<TN, CN>>;

export type SheetDataValues<
  TN extends TableNameSimple,
  CS extends ColumnName<TN> = ColumnName<TN>,
> = {
  [CN in CS]: ColumnValue<TN, CN>;
};

// Every writable column but the generated id: the bag a complete append must fill.
export type SheetDataValuesAll<TN extends TableNameSimple> = SheetDataValues<
  TN,
  Exclude<ColumnNameFiltered<TN, ValueName, false>, "id">
>;

export function getSheetColumnNames<TN extends TableNameSimple>(
  tableName: TN,
): ColumnName<TN>[] {
  return Obj.keys(columnConfigs()[tableName]) as unknown as ColumnName<TN>[];
}

// columnConfig isn't actually very unique. The only unique
export function getColumnTraitByName<
  TN extends TableNameSimple,
  CN extends ColumnName<TN>,
  TK extends keyof ColumnConfigAt<TN, CN>,
>(tableName: TN, columnName: CN, key: TK): ColumnConfigAt<TN, CN>[TK] {
  if (key === "columnName") {
    return columnName as ColumnConfigAt<TN, CN>[TK];
  }
  return (columnConfigs()[tableName][columnName] as ColumnConfigAt<TN, CN>)[
    key
  ];
}

export type TableColumnConfigsById = KeyedMap<
  Record<string, ColumnConfigStored>,
  "columnId",
  "columnName"
>;

type ColumnConfigsByTableAndColId = Map<
  TableNameSimple,
  TableColumnConfigsById
>;
function makeColumnConfigsByTableAndColId(): ColumnConfigsByTableAndColId {
  return configTableNames().reduce((attrs, tableName) => {
    attrs.set(
      tableName,
      Obj.toKeyedMap(columnConfigs()[tableName], "columnId", "columnName"),
    );
    return attrs;
  }, new Map() as ColumnConfigsByTableAndColId);
}

const columnConfigsByTableAndColId = lazy(makeColumnConfigsByTableAndColId);

export function getColumnTraitById<TK extends keyof ColumnConfig>(
  tableName: TableNameSimple,
  columnId: string,
  key: TK,
): ColumnConfig[TK] {
  const colTraits = Val.assert(
    columnConfigsByTableAndColId().get(tableName)?.get(columnId),
    `column attributes for Table ${tableName}, columnId=${columnId}`,
  );
  return colTraits[key];
}
export function getTableColumnIds(
  tableName: TableNameSimple,
): MapIterator<string> {
  return Val.assert(
    columnConfigsByTableAndColId().get(tableName),
    `column attributes for Table ${tableName}`,
  ).keys();
}

export type MakeColumnFullName<
  TN extends TableNameSimple,
  CN extends ColumnName<TN>,
> = `${TN}${CodebaseNameDelimiter}${CN & string}`;

type ColumnConfigsFlat = FlattenTwoLevels<
  ColumnConfigs,
  CodebaseNameDelimiter,
  "tableName",
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
  ColumnConfigsFlat[FN]["tableName"];
export type ColumnNameOf<FN extends ColumnFullName> =
  ColumnConfigsFlat[FN]["columnName"] & ColumnName<SheetNameOf<FN>>;
export type ValueNameOf<FN extends ColumnFullName> =
  ColumnConfigsFlat[FN]["valueName"];
export type ValueOf<FN extends ColumnFullName> = Value<ValueNameOf<FN>>;
