import type { SyncRefusal } from "../syncRunResult";
import type { ColumnNamedAt } from "./NewColumnNames";
import type { RenamedColumnNames } from "./RenamedColumnNames";

// What a Table entry lets a rule or brace name: its columns, the ones its `newColumns` create, and its renames.
export interface NameableColumns {
  names: string[] | undefined;
  renamedColumns: RenamedColumnNames | undefined;
}

// Only an unknown name is checked against the run's renames before `unknown` refuses it.
export function namedColumnRefusals(
  at: ColumnNamedAt,
  { names, renamedColumns }: NameableColumns,
  unknown: () => SyncRefusal,
): SyncRefusal[] {
  if (names?.includes(at.column)) return [];
  const renamed = renamedColumns?.refusals(at) ?? [];
  if (renamed.length > 0) return renamed;
  return [unknown()];
}
