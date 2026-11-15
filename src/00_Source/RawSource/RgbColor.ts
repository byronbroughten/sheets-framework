export interface RgbColor {
  red?: number;
  green?: number;
  blue?: number;
  alpha?: number;
}

// Rule colours omit alpha: a defaulted alpha makes every content comparison fail.
export type RgbChannels = Pick<RgbColor, "red" | "green" | "blue">;

export type HexColor = `#${string}`;

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

export function rgbChannelsHex({ red, green, blue }: RgbChannels): HexColor {
  return `#${[red, green, blue].map(hexChannel).join("")}`;
}

// The inverse of rgbChannelsHex, for a colour an agent wrote as `#rrggbb`.
export function hexRgbChannels(hex: string): RgbChannels {
  if (!/^#[0-9a-fA-F]{6}$/.test(hex)) {
    throw new Error(`The colour "${hex}" is not six-digit hex, as #00ff00 is.`);
  }
  const [red, green, blue] = [1, 3, 5].map(
    (start) => parseInt(hex.slice(start, start + 2), 16) / 255,
  );
  return { red, green, blue };
}

// Google omits a zero channel, so an absent one is 0.
function hexChannel(value: number | undefined): string {
  return Math.round((value ?? 0) * 255)
    .toString(16)
    .padStart(2, "0");
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
