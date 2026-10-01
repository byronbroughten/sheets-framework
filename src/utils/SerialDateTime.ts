import { SerialDate, zonedDateParts } from "./SerialDate";

// A serial's fraction is wall-clock time in the spreadsheet's zone, so an instant needs that date's offset, DST included.
export const SerialDateTime = {
  sheetsEpochUtcMs: SerialDate.sheetsEpochUtcMs,
  msPerDay: SerialDate.msPerDay,
  // Wall-clock date and time fields of `instant` as seen in `tz`.
  wallClockParts(instant: Date, tz: string): Record<string, string> {
    return zonedDateParts(instant, tz);
  },
  // Local wall-clock timestamp, e.g. "2026-08-31 17:14:10".
  nowTimestamp(tz: string): string {
    const p = SerialDateTime.wallClockParts(new Date(), tz);
    return `${p.year}-${p.month}-${p.day} ${p.hour}:${p.minute}:${p.second}`;
  },
  // UTC offset in minutes of `tz` at `instant`, positive east of UTC; DST-aware via Intl's IANA data.
  getTzOffsetMinutes(instant: Date, tz: string): number {
    const parts = SerialDateTime.wallClockParts(instant, tz);
    const asIfUTC = Date.UTC(
      Number(parts.year),
      Number(parts.month) - 1,
      Number(parts.day),
      Number(parts.hour),
      Number(parts.minute),
      Number(parts.second),
    );
    return (asIfUTC - instant.getTime()) / 60000;
  },
  // Sheets serial (with fractional time) -> JS Date, resolved against `tz`.
  serialToDateTime(serial: number, tz: string): Date {
    const naiveMs =
      SerialDateTime.sheetsEpochUtcMs +
      Math.round(serial * SerialDateTime.msPerDay);
    const offsetMin = SerialDateTime.getTzOffsetMinutes(new Date(naiveMs), tz);
    return new Date(naiveMs - offsetMin * 60000);
  },
  // JS Date -> Sheets serial (with fractional time), resolved against `tz`.
  dateTimeToSerial(date: Date, tz: string): number {
    const offsetMin = SerialDateTime.getTzOffsetMinutes(date, tz);
    const localMs = date.getTime() + offsetMin * 60000;
    return (
      (localMs - SerialDateTime.sheetsEpochUtcMs) / SerialDateTime.msPerDay
    );
  },
  // Add whole days on wall-clock fields in `tz` (DST-safe).
  addDaysTz(date: Date, days: number, tz: string): Date {
    const offsetMin = SerialDateTime.getTzOffsetMinutes(date, tz);
    const local = new Date(date.getTime() + offsetMin * 60000);
    local.setUTCDate(local.getUTCDate() + days);
    const newOffsetMin = SerialDateTime.getTzOffsetMinutes(local, tz);
    return new Date(local.getTime() - newOffsetMin * 60000);
  },
  // Add whole months on wall-clock fields in `tz` (DST-safe).
  addMonthsTz(date: Date, months: number, tz: string): Date {
    const offsetMin = SerialDateTime.getTzOffsetMinutes(date, tz);
    const local = new Date(date.getTime() + offsetMin * 60000);
    local.setUTCMonth(local.getUTCMonth() + months);
    const newOffsetMin = SerialDateTime.getTzOffsetMinutes(local, tz);
    return new Date(local.getTime() - newOffsetMin * 60000);
  },
};
