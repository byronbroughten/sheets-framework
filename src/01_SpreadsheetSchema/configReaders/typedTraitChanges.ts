import type {
  ColumnConfigsGeneric,
  ColumnConfigStored,
  ValueConfigBase,
  ValueConfigsBase,
} from "./makeConfigs";
import type { SyncSubject, SyncTypedTraitChange } from "./syncRunResult";

export interface TypedTraitConfigs {
  columnConfigs: ColumnConfigsGeneric;
  valueConfigs: ValueConfigsBase;
}

export interface TypedTraitChangesProps {
  previous: TypedTraitConfigs;
  regenerated: TypedTraitConfigs;
  // What the run itself sent: those columns aren't a person's change.
  sent: SyncSubject[];
}

/**
 * The value-kind changes between the previous and regenerated configs, per
 * column matched by ID, on Tables with API access in both. emptyValueAllowed
 * is typed too but isn't a sheet trait, so only an alteration changes it, and
 * a sent one is skipped like any other.
 */
export function typedTraitChanges({
  previous,
  regenerated,
  sent,
}: TypedTraitChangesProps): SyncTypedTraitChange[] {
  const sentColumns = new Set(sent.flatMap(sentColumnKey));
  return Object.entries(regenerated.columnConfigs).flatMap(([table, columns]) =>
    Object.entries(columns).flatMap(([column, config]) => {
      if (sentColumns.has(columnKey(table, column))) return [];
      const before = previousColumn(previous.columnConfigs[table], config);
      if (before === undefined) return [];
      return columnChanges({
        table,
        column,
        before: { config: before, valueConfigs: previous.valueConfigs },
        after: { config, valueConfigs: regenerated.valueConfigs },
      });
    }),
  );
}

interface ColumnSide {
  config: ColumnConfigStored;
  valueConfigs: ValueConfigsBase;
}

interface ColumnChangesProps {
  table: string;
  column: string;
  before: ColumnSide;
  after: ColumnSide;
}

function columnChanges({
  table,
  column,
  before,
  after,
}: ColumnChangesProps): SyncTypedTraitChange[] {
  const traits = [
    {
      trait: "columnType",
      from: before.config.columnType.type ?? "untyped",
      to: after.config.columnType.type ?? "untyped",
    },
    {
      trait: "valueName",
      from: before.config.valueName,
      to: after.config.valueName,
    },
    { trait: "options", from: optionsOf(before), to: optionsOf(after) },
  ];
  const [columnType, valueName, options] = traits;
  // A column type change usually renames the value too, so one line says it.
  const reported =
    columnType.from === columnType.to ? [valueName, options] : [columnType];
  return reported
    .filter(({ from, to }) => from !== to)
    .map((traitChange) => ({ table, column, ...traitChange }));
}

function previousColumn(
  columns: ColumnConfigsGeneric[string] | undefined,
  { columnId }: ColumnConfigStored,
): ColumnConfigStored | undefined {
  return Object.values(columns ?? {}).find(
    (config) => config.columnId === columnId,
  );
}

function optionsOf({ config, valueConfigs }: ColumnSide): string {
  const valueConfig: ValueConfigBase | undefined =
    valueConfigs[config.valueName];
  if (valueConfig === undefined) return "none";
  if (Array.isArray(valueConfig)) return valueConfig.join(", ");
  return "live reference";
}

function columnKey(table: string, column: string): string {
  return `${table}.${column}`;
}

function sentColumnKey(subject: SyncSubject): string[] {
  if (!("table" in subject) || subject.column === undefined) return [];
  return [columnKey(subject.table, subject.column)];
}
