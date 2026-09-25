export interface RgbColor {
  red?: number;
  green?: number;
  blue?: number;
  alpha?: number;
}

// Rule colours omit alpha: a defaulted alpha makes every content comparison fail.
export type RgbChannels = Pick<RgbColor, "red" | "green" | "blue">;

export function rgbChannelsEqual(
  left?: RgbChannels,
  right?: RgbChannels,
): boolean {
  if (left === undefined && right === undefined) return true;
  if (left === undefined || right === undefined) return false;
  return (
    quantizedChannel(left.red) === quantizedChannel(right.red) &&
    quantizedChannel(left.green) === quantizedChannel(right.green) &&
    quantizedChannel(left.blue) === quantizedChannel(right.blue)
  );
}

export function quantizeRgbChannels(color: RgbChannels): RgbChannels {
  return {
    ...(color.red !== undefined ? { red: quantizeRgbChannel(color.red) } : {}),
    ...(color.green !== undefined
      ? { green: quantizeRgbChannel(color.green) }
      : {}),
    ...(color.blue !== undefined
      ? { blue: quantizeRgbChannel(color.blue) }
      : {}),
  };
}

export function quantizeRgbChannel(value: number): number {
  return Math.round(value * 255) / 255;
}

function quantizedChannel(value: number | undefined): number | undefined {
  if (value === undefined) return undefined;
  return Math.round(value * 255);
}
