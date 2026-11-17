import {
  type BooleanCondition,
  type ConditionalFormat,
  conditionLiteral,
  isModelledConditionType,
  type ModelableConditionalFormatRule,
} from "../../00_Source/RawSource/ConditionalFormat";
import type { GridRangeProps } from "../../00_Source/RawSource/RawSource";
import { hexRgbChannels } from "../../00_Source/RawSource/RgbColor";
import { SheetIndex } from "../../00_Source/RawSource/SheetIndex";
import { BraceFormula, type RuleLayout } from "./BraceFormula";
import type { TableRuleRecord } from "./RuleRecords";
import type { WideRule, WideRuleFormat } from "./SyncPlanner/wideAlterations";

// A new column by the name the regen gives it, placed before `before` or at the Table's end.
export interface LayoutNewColumn {
  column: string;
  before?: string;
}

export interface PlannedLayoutProps {
  layout: RuleLayout;
  deletedColumns: string[];
  newColumns: LayoutNewColumn[];
}

export interface IntendedRuleProps {
  tableKey: string;
  sheetGid: number;
  layout: RuleLayout;
  rule: WideRule;
}

// An agent's intended rule, compiled to the rule Sheets holds; RuleRecords reads one back the other way.
export const IntendedRule = {
  // The Table's columns once the run's deletes and inserts land, in entry order as the inserts place them.
  plannedLayout({
    layout,
    deletedColumns,
    newColumns,
  }: PlannedLayoutProps): RuleLayout {
    const columnNames = layout.columnNamesLeftToRight.filter(
      (columnName) =>
        columnName === undefined || !deletedColumns.includes(columnName),
    );
    for (const { column, before } of newColumns) {
      const beforeOffset =
        before === undefined ? -1 : columnNames.indexOf(before);
      if (beforeOffset === -1) {
        columnNames.push(column);
        continue;
      }
      columnNames.splice(beforeOffset, 0, column);
    }
    return { ...layout, columnNamesLeftToRight: columnNames };
  },
  compile({
    tableKey,
    sheetGid,
    layout,
    rule,
  }: IntendedRuleProps): ModelableConditionalFormatRule {
    return {
      kind: "boolean",
      ranges: bodyRanges(sheetGid, layout, columnOffsets(layout, rule)),
      condition: compiledCondition({ tableKey, layout, rule }),
      format: compiledFormat(rule.format),
    };
  },
  // A recorded rule as an agent would write it, so a re-span compiles it as a replace does.
  ofRecord({ columns, condition, format }: TableRuleRecord): WideRule {
    return {
      columns,
      condition: {
        type: condition.type,
        values: [wideConditionValue(condition)],
      },
      format,
    };
  },
};

function wideConditionValue(condition: BooleanCondition): string {
  if (condition.type === "CUSTOM_FORMULA") return condition.formula;
  return conditionLiteral.fromCellValue(condition.value);
}

function columnOffsets(layout: RuleLayout, { columns }: WideRule): number[] {
  const names = layout.columnNamesLeftToRight;
  if (columns === "all") return names.map((_, offset) => offset);
  const offsets = columns.map((columnName) => {
    const offset = names.indexOf(columnName);
    if (offset === -1) {
      throw new Error(
        `A rule's column ${columnName} names no column of its Table once the run's column changes land.`,
      );
    }
    return offset;
  });
  return [...new Set(offsets)].sort((first, second) => first - second);
}

// One range per run of side-by-side columns, each spanning the body rows.
function bodyRanges(
  sheetGid: number,
  layout: RuleLayout,
  offsets: number[],
): GridRangeProps[] {
  const firstCol = layout.headerTopLeft.colIndex;
  const runs = offsets.reduce<{ start: number; end: number }[]>(
    (found, offset) => {
      const last = found[found.length - 1];
      if (last !== undefined && last.end === offset) {
        last.end = offset + 1;
        return found;
      }
      found.push({ start: offset, end: offset + 1 });
      return found;
    },
    [],
  );
  return runs.map(({ start, end }) => ({
    sheetId: sheetGid,
    startRowIndex: SheetIndex.row(layout.headerTopLeft.rowIndex + 1),
    endRowIndex: layout.endRowIndex,
    startColumnIndex: SheetIndex.col(firstCol + start),
    endColumnIndex: SheetIndex.col(firstCol + end),
  }));
}

function compiledCondition({
  tableKey,
  layout,
  rule,
}: Omit<IntendedRuleProps, "sheetGid">): BooleanCondition {
  const { type, values } = rule.condition;
  const value = values?.[0];
  if (!isModelledConditionType(type)) {
    throw new Error(`A rule's condition type ${type} isn't one sync can send.`);
  }
  if (value === undefined) {
    throw new Error(`A rule's ${type} condition has no value.`);
  }
  if (type === "CUSTOM_FORMULA") {
    return {
      type,
      formula: BraceFormula.compileForRule({
        tableKey,
        formula: value,
        layout,
      }),
    };
  }
  return { type, value: conditionLiteral.toCellValue(value) };
}

function compiledFormat({
  backgroundColor,
  foregroundColor,
}: WideRuleFormat): ConditionalFormat {
  return {
    ...(backgroundColor !== undefined
      ? { backgroundColor: hexRgbChannels(backgroundColor) }
      : {}),
    ...(foregroundColor !== undefined
      ? { foregroundColor: hexRgbChannels(foregroundColor) }
      : {}),
  };
}
