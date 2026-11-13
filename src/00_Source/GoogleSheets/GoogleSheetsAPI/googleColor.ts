import {
  type TabColor,
  type ThemeColorName,
  themeColorNames,
} from "../../RawSource/RawSource";
import {
  quantizeRgbChannels,
  type RgbChannels,
  rgbChannelsHex,
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
    return rgbChannelsHex(style.rgbColor);
  },
  fromTabColor(color: TabColor): GoogleColorStyle {
    const themeColor = themeColorNames.find((known) => known === color);
    if (themeColor !== undefined) return { themeColor };
    const [red, green, blue] = [1, 3, 5].map(
      (start) => parseInt(color.slice(start, start + 2), 16) / 255,
    );
    return { rgbColor: { red, green, blue } };
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
