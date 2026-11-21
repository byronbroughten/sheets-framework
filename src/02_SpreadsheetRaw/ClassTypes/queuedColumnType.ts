import type {
  TableColumnPropertiesAdd,
  TableColumnValidationRule,
} from "../../00_Source/RawSource/RawSource.ts";
import type { ColumnTypeBaseline } from "../../01_SpreadsheetSchema/columnBaseline.ts";
import type { ColumnStateRaw, QueuedColumnType } from "./StateRaw.ts";

// A queued column type is a bare type, which keeps the column's fetched rule, or a baseline, which sends its own.
export const queuedColumnType = {
  // Undefined for a bare type.
  baselineOf(queued: QueuedColumnType): ColumnTypeBaseline | undefined {
    if (typeof queued === "string") return undefined;
    return queued;
  },
  // What the column's type is once the flush sends it; undefined is Automatic.
  typeOf(queued: QueuedColumnType): string | undefined {
    const baseline = queuedColumnType.baselineOf(queued);
    if (baseline === undefined) return queued as string;
    return baseline.type ?? undefined;
  },
  ruleOf(
    queued: QueuedColumnType | undefined,
  ): TableColumnValidationRule | undefined {
    if (queued === undefined) return undefined;
    return queuedColumnType.baselineOf(queued)?.rule ?? undefined;
  },
  applyTo(columnState: ColumnStateRaw, queued: QueuedColumnType): void {
    const baseline = queuedColumnType.baselineOf(queued);
    if (baseline === undefined) {
      columnState.columnType = queued as string;
      return;
    }
    delete columnState.columnType;
    delete columnState.validationConditionType;
    delete columnState.validationValues;
    if (baseline.type !== null) columnState.columnType = baseline.type;
    if (baseline.rule === null) return;
    columnState.validationConditionType = baseline.rule.condition;
    if (baseline.rule.values.length > 0) {
      columnState.validationValues = baseline.rule.values;
    }
  },
  // A null type or rule is left out, which leaves the column on Automatic or with no rule.
  properties({
    type,
    rule,
  }: ColumnTypeBaseline): Pick<
    TableColumnPropertiesAdd,
    "columnType" | "dataValidationRule"
  > {
    return {
      ...(type === null ? {} : { columnType: type }),
      ...(rule === null ? {} : { dataValidationRule: rule }),
    };
  },
};
