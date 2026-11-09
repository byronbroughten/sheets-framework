import {
  type TabColor,
  type ThemeColorName,
  themeColorNames,
} from "../../RawSource/RawSource";
import {
  quantizeRgbChannels,
  type RgbChannels,
  type RgbColor,
} from "../../RawSource/RgbColor";

type GoogleColor = GoogleAppsScript.Sheets.Schema.Color;
type GoogleColorStyle = GoogleAppsScript.Sheets.Schema.ColorStyle;

export const googleColor = {
  fromRgb(color: RgbColor): GoogleColor {
    return {
      red: color.red,
      green: color.green,
      blue: color.blue,
      alpha: color.alpha,
    };
  },
  toRgb(color: GoogleColor): RgbColor {
    return {
      red: color.red,
      green: color.green,
      blue: color.blue,
      alpha: color.alpha,
    };
  },
  fromRgbChannels(color: RgbChannels): GoogleColor {
    return {
      ...(color.red !== undefined ? { red: color.red } : {}),
      ...(color.green !== undefined ? { green: color.green } : {}),
      ...(color.blue !== undefined ? { blue: color.blue } : {}),
    };
  },
  toRgbChannels(color: GoogleColor): RgbChannels {
    return quantizeRgbChannels(color);
  },
  toTabColor(style: GoogleColorStyle | undefined): TabColor | undefined {
    if (style?.themeColor !== undefined) {
      return themeColorName(style.themeColor);
    }
    if (style?.rgbColor === undefined) return undefined;
    const { red, green, blue } = style.rgbColor;
    return `#${[red, green, blue].map(hexChannel).join("")}`;
  },
};

function themeColorName(themeColor: string): ThemeColorName {
  const name = themeColorNames.find((known) => known === themeColor);
  if (name === undefined) {
    throw new Error(
      `Theme colour "${themeColor}" is not one the framework knows.`,
    );
  }
  return name;
}

// Google omits a zero channel, so an absent one is 0.
function hexChannel(value: number | undefined): string {
  return Math.round((value ?? 0) * 255)
    .toString(16)
    .padStart(2, "0");
}
