import { SerialDate } from "./SerialDate";

// ---------------------------------------------------------------------
// DATE-TIME (has a time-of-day) — for later, once you have time columns
// ---------------------------------------------------------------------
// A Sheets serial's fractional part is LOCAL wall-clock time in the
// spreadsheet's timezone (e.g. Central). Resolving that to a real
// instant requires the actual UTC offset for that date, DST included.

export const Tim = {
  sheetsEpochUtcMs: SerialDate.sheetsEpochUtcMs,
  msPerDay: SerialDate.msPerDay,
  // Wall-clock date and time fields of `instant` as seen in `tz`.
  wallClockParts(instant: Date, tz: string): Record<string, string> {
    return new Intl.DateTimeFormat("en-US", {
      timeZone: tz,
      hourCycle: "h23",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
    })
      .formatToParts(instant)
      .reduce<Record<string, string>>((acc, part) => {
        acc[part.type] = part.value;
        return acc;
      }, {});
  },
  // Local wall-clock timestamp, e.g. "2026-08-31 17:14:10".
  nowTimestamp(tz: string): string {
    const p = this.wallClockParts(new Date(), tz);
    return `${p.year}-${p.month}-${p.day} ${p.hour}:${p.minute}:${p.second}`;
  },
  // UTC offset (in minutes) of `tz` at the real-world moment `instant`
  // represents. Positive = east of UTC. DST-aware via Intl + the IANA
  // tz database.
  getTzOffsetMinutes(instant: Date, tz: string): number {
    const parts = this.wallClockParts(instant, tz);
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
    const naiveMs = this.sheetsEpochUtcMs + Math.round(serial * this.msPerDay);
    const offsetMin = this.getTzOffsetMinutes(new Date(naiveMs), tz);
    return new Date(naiveMs - offsetMin * 60000);
  },
  // JS Date -> Sheets serial (with fractional time), resolved against `tz`.
  dateTimeToSerial(date: Date, tz: string): number {
    const offsetMin = this.getTzOffsetMinutes(date, tz);
    const localMs = date.getTime() + offsetMin * 60000;
    return (localMs - this.sheetsEpochUtcMs) / this.msPerDay;
  },
  // Add whole days on wall-clock fields in `tz` (DST-safe).
  addDaysTz(date: Date, days: number, tz: string): Date {
    const offsetMin = this.getTzOffsetMinutes(date, tz);
    const local = new Date(date.getTime() + offsetMin * 60000);
    local.setUTCDate(local.getUTCDate() + days);
    const newOffsetMin = this.getTzOffsetMinutes(local, tz);
    return new Date(local.getTime() - newOffsetMin * 60000);
  },
  // Add whole months on wall-clock fields in `tz` (DST-safe).
  addMonthsTz(date: Date, months: number, tz: string): Date {
    const offsetMin = this.getTzOffsetMinutes(date, tz);
    const local = new Date(date.getTime() + offsetMin * 60000);
    local.setUTCMonth(local.getUTCMonth() + months);
    const newOffsetMin = this.getTzOffsetMinutes(local, tz);
    return new Date(local.getTime() - newOffsetMin * 60000);
  },
};
