import type { CellValue } from "../CellValues/cellValues";
import { type GridRangeProps, rangeEqual } from "./RawSource";
import { type RgbChannels, rgbChannelsEqual } from "./RgbColor";

export const modelledConditionTypes = [
  "NUMBER_EQ",
  "NUMBER_NOT_EQ",
  "NUMBER_GREATER",
  "NUMBER_GREATER_THAN_EQ",
  "NUMBER_LESS",
  "NUMBER_LESS_THAN_EQ",
  "TEXT_EQ",
  "CUSTOM_FORMULA",
] as const;

export type ModelledConditionType = (typeof modelledConditionTypes)[number];
export type ComparisonConditionType = Exclude<
  ModelledConditionType,
  "CUSTOM_FORMULA"
>;

export type BooleanCondition =
  | { type: "CUSTOM_FORMULA"; formula: string }
  | { type: ComparisonConditionType; value: CellValue };

export interface ConditionalFormat {
  backgroundColor?: RgbChannels;
  foregroundColor?: RgbChannels;
}

export interface ConditionalFormatDeclaration {
  condition: BooleanCondition;
  format: ConditionalFormat;
}

export type ConditionalFormatRule =
  ModelableConditionalFormatRule | UnmodelableConditionalFormatRule;

export interface ModelableConditionalFormatRule extends ConditionalFormatDeclaration {
  kind: "boolean";
  ranges: GridRangeProps[];
}

export interface UnmodelableConditionalFormatRule {
  kind: "unmodelable";
  ranges: GridRangeProps[];
  index: number;
}

export function isModelledConditionType(
  type: string,
): type is ModelledConditionType {
  return modelledConditionTypes.some((modelledType) => modelledType === type);
}

export function conditionalFormatRulesEqual(
  left: ConditionalFormatRule,
  right: ConditionalFormatRule,
): boolean {
  if (left.kind !== "boolean" || right.kind !== "boolean") return false;
  return (
    rangesEqual(left.ranges, right.ranges) &&
    conditionsEqual(left.condition, right.condition) &&
    formatsEqual(left.format, right.format)
  );
}

function rangesEqual(left: GridRangeProps[], right: GridRangeProps[]): boolean {
  if (left.length !== right.length) return false;
  return left.every((range, index) => rangeEqual(range, right[index]));
}

function conditionsEqual(
  left: BooleanCondition,
  right: BooleanCondition,
): boolean {
  if (left.type === "CUSTOM_FORMULA" && right.type === "CUSTOM_FORMULA") {
    return left.formula === right.formula;
  }
  if (left.type === "CUSTOM_FORMULA" || right.type === "CUSTOM_FORMULA") {
    return false;
  }
  return left.type === right.type && left.value === right.value;
}

function formatsEqual(
  left: ConditionalFormat,
  right: ConditionalFormat,
): boolean {
  return (
    rgbChannelsEqual(left.backgroundColor, right.backgroundColor) &&
    rgbChannelsEqual(left.foregroundColor, right.foregroundColor)
  );
}
