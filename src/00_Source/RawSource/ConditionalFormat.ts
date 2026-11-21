import type { CellValue } from "../CellValues/cellValues";
import { type GridRangeProps, rangeEqual } from "./RawSource";
import { type RgbChannels, rgbChannelsEqual } from "./RgbColor";

const numberConditionTypes = [
  "NUMBER_EQ",
  "NUMBER_NOT_EQ",
  "NUMBER_GREATER",
  "NUMBER_GREATER_THAN_EQ",
  "NUMBER_LESS",
  "NUMBER_LESS_THAN_EQ",
] as const;

export const modelledConditionTypes = [
  ...numberConditionTypes,
  "TEXT_EQ",
  "CUSTOM_FORMULA",
] as const;

export type NumberConditionType = (typeof numberConditionTypes)[number];
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
  rawContent: string;
}

// A comparison condition's value as Sheets' userEnteredValue text holds it.
export const conditionLiteral = {
  fromCellValue(value: CellValue): string {
    if (value === true) return "TRUE";
    if (value === false) return "FALSE";
    return String(value);
  },
  toCellValue(text: string): CellValue {
    if (text === "TRUE") return true;
    if (text === "FALSE") return false;
    if (text !== "" && Number(text).toString() === text) return Number(text);
    return text;
  },
};

export function isModelledConditionType(
  type: string,
): type is ModelledConditionType {
  return modelledConditionTypes.some((modelledType) => modelledType === type);
}

export function isNumberConditionType(
  type: string,
): type is NumberConditionType {
  return numberConditionTypes.some((numberType) => numberType === type);
}

// The GUI saves a NUMBER_* rule holding text that updateConditionalFormatRule refuses; observed: docs/sheets-behaviour.md#reading-conditional-formats and #writing-conditional-formats.
export function numberConditionHoldsText(
  type: string,
  userEnteredValue: string,
): boolean {
  if (!isNumberConditionType(type)) return false;
  const trimmed = userEnteredValue.trim();
  return trimmed === "" || !Number.isFinite(Number(trimmed));
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
