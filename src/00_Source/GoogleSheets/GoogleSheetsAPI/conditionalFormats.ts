import type { CellValue } from "../../CellValues/cellValues";
import {
  type BooleanCondition,
  type ConditionalFormat,
  type ConditionalFormatRule,
  isModelledConditionType,
  type ModelableConditionalFormatRule,
  type ModelledConditionType,
  quantizeConditionalFormat,
} from "../../RawSource/ConditionalFormat";
import type { GridRangeProps } from "../../RawSource/RawSource";
import { googleColor } from "./googleColor";
import { googleGrid } from "./gridSnapshots";

type GoogleColor = GoogleAppsScript.Sheets.Schema.Color;
type GoogleConditionalFormatRule =
  GoogleAppsScript.Sheets.Schema.ConditionalFormatRule;
type GoogleCellFormat = GoogleAppsScript.Sheets.Schema.CellFormat;

export const googleConditionalFormatRule = {
  fromModelable(
    rule: ModelableConditionalFormatRule,
  ): GoogleConditionalFormatRule {
    return {
      ranges: rule.ranges,
      booleanRule: {
        condition: googleBooleanCondition.fromCondition(rule.condition),
        format: googleCellFormat.fromConditionalFormat(rule.format),
      },
    };
  },
  toConditionalFormatRule(
    rule: GoogleConditionalFormatRule,
    index: number,
  ): ConditionalFormatRule {
    const ranges = (rule.ranges ?? []).map(googleGrid.toGridRangeProps);
    const modelled = toModelableRule(rule, ranges);
    if (modelled === null) {
      return { kind: "unmodelable", ranges, index };
    }
    return modelled;
  },
};

const googleBooleanCondition = {
  fromCondition(
    condition: BooleanCondition,
  ): GoogleAppsScript.Sheets.Schema.BooleanCondition {
    if (condition.type === "CUSTOM_FORMULA") {
      return {
        type: "CUSTOM_FORMULA",
        values: [{ userEnteredValue: condition.formula }],
      };
    }
    return {
      type: condition.type,
      values: [
        { userEnteredValue: conditionLiteral.fromCellValue(condition.value) },
      ],
    };
  },
  toCondition(
    type: ModelledConditionType,
    userEnteredValue: string,
  ): BooleanCondition {
    if (type === "CUSTOM_FORMULA") {
      return { type, formula: userEnteredValue };
    }
    return {
      type,
      value: conditionLiteral.toCellValue(userEnteredValue),
    };
  },
};

const googleCellFormat = {
  fromConditionalFormat(format: ConditionalFormat): GoogleCellFormat {
    return {
      ...(format.backgroundColor !== undefined
        ? {
            backgroundColor: googleColor.fromRgbChannels(
              format.backgroundColor,
            ),
          }
        : {}),
      ...(format.foregroundColor !== undefined
        ? {
            textFormat: {
              foregroundColor: googleColor.fromRgbChannels(
                format.foregroundColor,
              ),
            },
          }
        : {}),
    };
  },
  toConditionalFormat(format: GoogleCellFormat | undefined): ConditionalFormat {
    const backgroundColor = format?.backgroundColor;
    const foregroundColor = format?.textFormat?.foregroundColor;
    return {
      ...(backgroundColor !== undefined
        ? { backgroundColor: googleColor.toRgbChannels(backgroundColor) }
        : {}),
      ...(foregroundColor !== undefined
        ? { foregroundColor: googleColor.toRgbChannels(foregroundColor) }
        : {}),
    };
  },
};

const conditionLiteral = {
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

function toModelableRule(
  rule: GoogleConditionalFormatRule,
  ranges: GridRangeProps[],
): ModelableConditionalFormatRule | null {
  if (rule.gradientRule !== undefined) return null;
  const booleanRule = rule.booleanRule;
  if (booleanRule === undefined) return null;
  if (cellFormatHasUnmodelledFields(booleanRule.format)) return null;
  const type = booleanRule.condition?.type;
  if (type === undefined || !isModelledConditionType(type)) return null;
  const userEnteredValue = booleanRule.condition?.values?.[0]?.userEnteredValue;
  if (userEnteredValue === undefined) return null;
  const condition = googleBooleanCondition.toCondition(type, userEnteredValue);
  return {
    kind: "boolean",
    ranges,
    condition,
    format: quantizeConditionalFormat(
      googleCellFormat.toConditionalFormat(booleanRule.format),
    ),
  };
}

function cellFormatHasUnmodelledFields(
  format: GoogleCellFormat | undefined,
): boolean {
  if (format === undefined) return false;
  if (
    objectHasDefinedKeysBesides(format, [
      "backgroundColor",
      "backgroundColorStyle",
      "textFormat",
    ]) ||
    !colorStyleMirrors(format.backgroundColorStyle, format.backgroundColor)
  ) {
    return true;
  }
  const textFormat = format.textFormat;
  if (textFormat === undefined) return false;
  return (
    objectHasDefinedKeysBesides(textFormat, [
      "foregroundColor",
      "foregroundColorStyle",
    ]) ||
    !colorStyleMirrors(
      textFormat.foregroundColorStyle,
      textFormat.foregroundColor,
    )
  );
}

function objectHasDefinedKeysBesides(
  object: object,
  allowed: string[],
): boolean {
  return Object.entries(object).some(
    ([key, value]) => value !== undefined && !allowed.includes(key),
  );
}

// Google echoes each colour as a style; only an rgb copy of the plain colour is modelled.
function colorStyleMirrors(
  style: GoogleAppsScript.Sheets.Schema.ColorStyle | undefined,
  color: GoogleColor | undefined,
): boolean {
  if (style === undefined) return true;
  const rgb = style.rgbColor;
  if (style.themeColor !== undefined || rgb === undefined) return false;
  if (color === undefined) return false;
  return (["red", "green", "blue", "alpha"] as const).every(
    (channel) => (rgb[channel] ?? 0) === (color[channel] ?? 0),
  );
}
