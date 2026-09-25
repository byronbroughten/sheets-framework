import {
  quantizeRgbChannels,
  type RgbChannels,
  type RgbColor,
} from "../../RawSource/RgbColor";

type GoogleColor = GoogleAppsScript.Sheets.Schema.Color;

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
    return quantizeRgbChannels({
      ...(color.red !== undefined ? { red: color.red } : {}),
      ...(color.green !== undefined ? { green: color.green } : {}),
      ...(color.blue !== undefined ? { blue: color.blue } : {}),
    });
  },
};
