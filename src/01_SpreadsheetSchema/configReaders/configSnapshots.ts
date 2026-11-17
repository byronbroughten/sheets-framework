import type { IdNameMapGeneric } from "./idNameMapTypes";
import type {
  ColumnConfigsGeneric,
  TableConfigsBase,
  ValueConfigsBase,
} from "./makeConfigs";

// The configs on one side of a regen.
export interface ConfigSnapshot {
  idNameMap: IdNameMapGeneric;
  tableConfigs: TableConfigsBase;
  columnConfigs: ColumnConfigsGeneric;
  valueConfigs: ValueConfigsBase;
}

// The configs a regen started from and the ones it emits.
export interface ConfigSnapshots {
  previous: ConfigSnapshot;
  regenerated: ConfigSnapshot;
}
