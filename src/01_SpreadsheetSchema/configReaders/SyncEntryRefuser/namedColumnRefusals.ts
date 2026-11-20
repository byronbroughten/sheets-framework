import type { SyncRefusal } from "../syncRunResult";
import type { ColumnNamedAt, NewColumnNames } from "./NewColumnNames";
import type { RenamedColumnNames } from "./RenamedColumnNames";

// What a Table entry lets a rule or brace name: its columns, the ones its `newColumns` create, and its renames.
export interface NameableColumns {
  names: string[] | undefined;
  newColumns: NewColumnNames | undefined;
  renamedColumns: RenamedColumnNames | undefined;
}

// Ambiguity outranks a known name, and only an unknown one is checked against the run's renames before `unknown` refuses it.
export function namedColumnRefusals(
  at: ColumnNamedAt,
  { names, newColumns, renamedColumns }: NameableColumns,
  unknown: () => SyncRefusal,
): SyncRefusal[] {
  const ambiguous = newColumns?.ambiguousRefusals(at) ?? [];
  if (ambiguous.length > 0) return ambiguous;
  if (names?.includes(at.column)) return [];
  const renamed = renamedColumns?.refusals(at) ?? [];
  if (renamed.length > 0) return renamed;
  return [unknown()];
}
