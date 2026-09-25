import { describe, expect, it } from "vitest";

import { SerialDate } from "./SerialDate";

function ymd(year: number, month: number, day: number): SerialDate {
  return SerialDate.fromYmd({ year, month, day });
}

describe("SerialDate.validate", () => {
  it("returns the serial it was given for a whole number of days", () => {
    expect(SerialDate.validate(45292)).toBe(45292);
    expect(SerialDate.validate(0)).toBe(0);
    expect(SerialDate.validate(-1)).toBe(-1);
  });

  it("throws for a serial carrying a time of day", () => {
    expect(() => SerialDate.validate(45292.5)).toThrowError(/whole-day/);
  });

  it("throws for a non-number, NaN, or an infinity", () => {
    expect(() => SerialDate.validate("45292")).toThrowError(/whole-day/);
    expect(() => SerialDate.validate(NaN)).toThrowError(/whole-day/);
    expect(() => SerialDate.validate(Infinity)).toThrowError(/whole-day/);
    expect(() => SerialDate.validate(new Date())).toThrowError(/whole-day/);
  });
});

describe("SerialDate.isSerial", () => {
  it("narrows a whole number and rejects everything else", () => {
    expect(SerialDate.isSerial(45292)).toBe(true);
    expect(SerialDate.isSerial(45292.5)).toBe(false);
    expect(SerialDate.isSerial("45292")).toBe(false);
    expect(SerialDate.isSerial(new Date())).toBe(false);
  });
});

describe("SerialDate.fromYmd / SerialDate.toYmd", () => {
  it("counts days from the Sheets epoch", () => {
    expect(ymd(1899, 12, 30)).toBe(0);
    expect(ymd(1970, 1, 1)).toBe(25569);
  });

  it("round-trips a civil date, with the month as 1-12", () => {
    expect(SerialDate.toYmd(ymd(2024, 1, 31))).toEqual({
      year: 2024,
      month: 1,
      day: 31,
    });
    expect(SerialDate.toYmd(ymd(2024, 12, 1))).toEqual({
      year: 2024,
      month: 12,
      day: 1,
    });
  });

  it("throws on a day that does not exist rather than overflowing", () => {
    expect(() => ymd(2023, 2, 29)).toThrowError(/2023-2-29/);
    expect(() => ymd(2024, 4, 31)).toThrowError(/not a real date/);
    expect(() => ymd(2024, 0, 1)).toThrowError(/not a real date/);
    expect(() => ymd(2024, 13, 1)).toThrowError(/not a real date/);
    expect(() => ymd(2024, 1, 0)).toThrowError(/not a real date/);
  });

  it("throws on a non-integer part", () => {
    expect(() => ymd(2024, 1, 1.5)).toThrowError(/not a real date/);
  });

  it("accepts 29 February in a leap year", () => {
    expect(SerialDate.toYmd(ymd(2024, 2, 29)).day).toBe(29);
  });
});

describe("SerialDate.toDayMonthYear", () => {
  it("writes the civil day as day, abbreviated English month, year", () => {
    expect(SerialDate.toDayMonthYear(ymd(2026, 9, 12))).toBe("12 Sep 2026");
    expect(SerialDate.toDayMonthYear(ymd(2023, 3, 5))).toBe("5 Mar 2023");
  });
});

describe("SerialDate.fromInstant", () => {
  function inZone(isoInstant: string, timeZone: string): SerialDate {
    return SerialDate.fromInstant(new Date(isoInstant), timeZone);
  }

  it("is the civil day of the instant in the given zone, not in UTC", () => {
    const instant = "2024-03-15T02:30:00Z";

    expect(inZone(instant, "UTC")).toBe(ymd(2024, 3, 15));
    expect(inZone(instant, "America/Chicago")).toBe(ymd(2024, 3, 14));
    expect(inZone(instant, "Asia/Tokyo")).toBe(ymd(2024, 3, 15));
  });

  it("follows zones at the far ends of the offset range and on half hours", () => {
    const instant = "2024-06-15T10:15:00Z";

    expect(inZone(instant, "Pacific/Pago_Pago")).toBe(ymd(2024, 6, 14));
    expect(inZone(instant, "Pacific/Kiritimati")).toBe(ymd(2024, 6, 16));
    expect(inZone("2024-06-15T18:29:59Z", "Asia/Kolkata")).toBe(
      ymd(2024, 6, 15),
    );
    expect(inZone("2024-06-15T18:30:00Z", "Asia/Kolkata")).toBe(
      ymd(2024, 6, 16),
    );
  });

  it("rolls the month and year with the zone's midnight", () => {
    expect(inZone("2024-12-31T15:00:00Z", "Asia/Tokyo")).toBe(ymd(2025, 1, 1));
    expect(inZone("2025-01-01T05:59:59Z", "America/Chicago")).toBe(
      ymd(2024, 12, 31),
    );
  });

  it("moves midnight with the spring-forward offset", () => {
    expect(inZone("2024-03-10T05:59:59Z", "America/Chicago")).toBe(
      ymd(2024, 3, 9),
    );
    expect(inZone("2024-03-10T06:00:00Z", "America/Chicago")).toBe(
      ymd(2024, 3, 10),
    );
    expect(inZone("2024-03-10T08:00:00Z", "America/Chicago")).toBe(
      ymd(2024, 3, 10),
    );
    expect(inZone("2024-03-11T04:59:59Z", "America/Chicago")).toBe(
      ymd(2024, 3, 10),
    );
    expect(inZone("2024-03-11T05:00:00Z", "America/Chicago")).toBe(
      ymd(2024, 3, 11),
    );
  });

  it("keeps both passes of the repeated fall-back hour on one day", () => {
    expect(inZone("2024-11-03T06:30:00Z", "America/Chicago")).toBe(
      ymd(2024, 11, 3),
    );
    expect(inZone("2024-11-03T07:30:00Z", "America/Chicago")).toBe(
      ymd(2024, 11, 3),
    );
    expect(inZone("2024-11-04T05:59:59Z", "America/Chicago")).toBe(
      ymd(2024, 11, 3),
    );
    expect(inZone("2024-11-04T06:00:00Z", "America/Chicago")).toBe(
      ymd(2024, 11, 4),
    );
  });

  it("skips a civil day the zone itself skipped", () => {
    expect(inZone("2011-12-30T09:59:59Z", "Pacific/Apia")).toBe(
      ymd(2011, 12, 29),
    );
    expect(inZone("2011-12-30T10:00:00Z", "Pacific/Apia")).toBe(
      ymd(2011, 12, 31),
    );
  });
});

describe("SerialDate.addDays / SerialDate.dayBefore", () => {
  it("adds whole days as integer arithmetic", () => {
    const start = ymd(2024, 2, 28);

    expect(SerialDate.addDays(start, 1)).toBe(start + 1);
    expect(SerialDate.addDays(start, 2)).toBe(ymd(2024, 3, 1));
    expect(SerialDate.addDays(start, -28)).toBe(ymd(2024, 1, 31));
  });

  it("names the day before a start date", () => {
    expect(SerialDate.dayBefore(ymd(2024, 3, 1))).toBe(ymd(2024, 2, 29));
  });
});

describe("SerialDate.addMonths", () => {
  it("keeps the day of month when the target month has it", () => {
    expect(SerialDate.addMonths(ymd(2024, 1, 15), 1)).toBe(ymd(2024, 2, 15));
    expect(SerialDate.addMonths(ymd(2024, 1, 15), 12)).toBe(ymd(2025, 1, 15));
    expect(SerialDate.addMonths(ymd(2024, 3, 15), -3)).toBe(ymd(2023, 12, 15));
  });

  it("clips to the last day of the target month instead of overflowing", () => {
    expect(SerialDate.addMonths(ymd(2024, 1, 31), 1)).toBe(ymd(2024, 2, 29));
    expect(SerialDate.addMonths(ymd(2023, 1, 31), 1)).toBe(ymd(2023, 2, 28));
    expect(SerialDate.addMonths(ymd(2024, 5, 31), 1)).toBe(ymd(2024, 6, 30));
  });
});

describe("SerialDate.isSameOrAfter / isSameOrBefore", () => {
  it("counts the same day as both same-or-after and same-or-before", () => {
    const date = ymd(2024, 6, 1);

    expect(SerialDate.isSameOrAfter(date, date)).toBe(true);
    expect(SerialDate.isSameOrBefore(date, date)).toBe(true);
    expect(SerialDate.isSameOrAfter(date, ymd(2024, 5, 31))).toBe(true);
    expect(SerialDate.isSameOrBefore(date, ymd(2024, 5, 31))).toBe(false);
  });
});

describe("SerialDate.isOnOrBetween", () => {
  const startDate = ymd(2024, 6, 1);
  const endDate = ymd(2024, 6, 30);

  it("includes both ends of the window", () => {
    expect(
      SerialDate.isOnOrBetween({ date: startDate, startDate, endDate }),
    ).toBe(true);
    expect(
      SerialDate.isOnOrBetween({ date: endDate, startDate, endDate }),
    ).toBe(true);
    expect(
      SerialDate.isOnOrBetween({ date: ymd(2024, 6, 15), startDate, endDate }),
    ).toBe(true);
  });

  it("excludes a date outside the window on either side", () => {
    expect(
      SerialDate.isOnOrBetween({ date: ymd(2024, 5, 31), startDate, endDate }),
    ).toBe(false);
    expect(
      SerialDate.isOnOrBetween({ date: ymd(2024, 7, 1), startDate, endDate }),
    ).toBe(false);
  });

  it("throws when the start is after the end", () => {
    expect(() =>
      SerialDate.isOnOrBetween({
        date: startDate,
        startDate: endDate,
        endDate: startDate,
      }),
    ).toThrowError(/after end date/);
  });
});

describe("SerialDate.monthYear / isInMonthAndYear / monthYearsOnAndBetween", () => {
  it("reads January as month 1", () => {
    expect(SerialDate.monthYear(ymd(2024, 1, 15))).toEqual({
      month: 1,
      year: 2024,
    });
  });

  it("counts a date as in its own month and year", () => {
    const date = ymd(2024, 1, 15);

    expect(SerialDate.isInMonthAndYear(date, { month: 1, year: 2024 })).toBe(
      true,
    );
    expect(SerialDate.isInMonthAndYear(date, { month: 2, year: 2024 })).toBe(
      false,
    );
    expect(SerialDate.isInMonthAndYear(date, { month: 1, year: 2025 })).toBe(
      false,
    );
  });

  it("walks every month from the start through the end, inclusive", () => {
    expect(
      SerialDate.monthYearsOnAndBetween({
        startMonthYear: { month: 11, year: 2023 },
        endMonthYear: { month: 2, year: 2024 },
      }),
    ).toEqual([
      { month: 11, year: 2023 },
      { month: 12, year: 2023 },
      { month: 1, year: 2024 },
      { month: 2, year: 2024 },
    ]);
  });
});

describe("SerialDate month bounds", () => {
  it("takes a serial for firstDayOfMonth and lastDayOfMonth", () => {
    const date = ymd(2024, 2, 15);

    expect(SerialDate.firstDayOfMonth(date)).toBe(ymd(2024, 2, 1));
    expect(SerialDate.lastDayOfMonth(date)).toBe(ymd(2024, 2, 29));
    expect(SerialDate.firstAndLastDayOfMonth(date)).toEqual({
      firstOfMonth: ymd(2024, 2, 1),
      lastOfMonth: ymd(2024, 2, 29),
    });
  });

  it("takes a month and year for the MonthYear forms", () => {
    const monthYear = { month: 2, year: 2023 };

    expect(SerialDate.firstDayOfMonthYear(monthYear)).toBe(ymd(2023, 2, 1));
    expect(SerialDate.lastDayOfMonthYear(monthYear)).toBe(ymd(2023, 2, 28));
    expect(SerialDate.firstAndLastDayOfMonthYear(monthYear)).toEqual({
      firstOfMonth: ymd(2023, 2, 1),
      lastOfMonth: ymd(2023, 2, 28),
    });
  });

  it("gives the first day of the following month, unlike the other two", () => {
    const date = ymd(2024, 12, 31);

    expect(SerialDate.firstDayOfNextMonth(date)).toBe(ymd(2025, 1, 1));
    expect(SerialDate.firstDayOfMonth(date)).toBe(ymd(2024, 12, 1));
    expect(SerialDate.firstDayOfMonthYear({ month: 12, year: 2024 })).toBe(
      ymd(2024, 12, 1),
    );
    expect(SerialDate.addMonths(date, 1)).toBe(ymd(2025, 1, 31));
  });
});

describe("SerialDate.proratedMonthlyProportion", () => {
  it("gives exactly 1 for a range covering a whole month", () => {
    expect(
      SerialDate.proratedMonthlyProportion({
        startDate: ymd(2024, 1, 1),
        endDate: ymd(2024, 1, 31),
      }),
    ).toBe(1);
  });

  it("counts inclusive days over the month length when the range starts mid-month", () => {
    expect(
      SerialDate.proratedMonthlyProportion({
        startDate: ymd(2024, 1, 16),
        endDate: ymd(2024, 1, 31),
      }),
    ).toBe(16 / 31);
  });

  it("counts inclusive days over the month length when the range ends mid-month", () => {
    expect(
      SerialDate.proratedMonthlyProportion({
        startDate: ymd(2024, 2, 1),
        endDate: ymd(2024, 2, 15),
      }),
    ).toBe(15 / 29);
  });

  it("divides by 29 in a leap February and 28 in a common one", () => {
    expect(
      SerialDate.proratedMonthlyProportion({
        startDate: ymd(2024, 2, 1),
        endDate: ymd(2024, 2, 10),
      }),
    ).toBe(10 / 29);
    expect(
      SerialDate.proratedMonthlyProportion({
        startDate: ymd(2023, 2, 1),
        endDate: ymd(2023, 2, 10),
      }),
    ).toBe(10 / 28);
  });

  it("counts a single day as one day of the month", () => {
    expect(
      SerialDate.proratedMonthlyProportion({
        startDate: ymd(2024, 4, 7),
        endDate: ymd(2024, 4, 7),
      }),
    ).toBe(1 / 30);
  });

  it("throws naming both months when the range spans two of them", () => {
    expect(() =>
      SerialDate.proratedMonthlyProportion({
        startDate: ymd(2024, 1, 16),
        endDate: ymd(2024, 3, 15),
      }),
    ).toThrowError(/2024-1.*2024-3/);
  });

  it("throws when the end is before the start", () => {
    expect(() =>
      SerialDate.proratedMonthlyProportion({
        startDate: ymd(2024, 1, 16),
        endDate: ymd(2024, 1, 15),
      }),
    ).toThrowError(/after end date/);
  });
});

describe("SerialDate.proratedMonthlyAmount", () => {
  it("takes the amount first and returns the unrounded product", () => {
    expect(
      SerialDate.proratedMonthlyAmount(3100, {
        startDate: ymd(2024, 1, 16),
        endDate: ymd(2024, 1, 31),
      }),
    ).toBe(3100 * (16 / 31));
  });

  it("returns the whole amount for a whole month", () => {
    expect(
      SerialDate.proratedMonthlyAmount(1000, {
        startDate: ymd(2024, 1, 1),
        endDate: ymd(2024, 1, 31),
      }),
    ).toBe(1000);
  });
});

describe("SerialDate.monthRanges", () => {
  it("gives one entry for a term inside one month", () => {
    expect(
      SerialDate.monthRanges({
        startDate: ymd(2024, 1, 5),
        endDate: ymd(2024, 1, 20),
      }),
    ).toEqual([
      {
        month: 1,
        year: 2024,
        startDate: ymd(2024, 1, 5),
        endDate: ymd(2024, 1, 20),
      },
    ]);
  });

  it("gives one entry of one day for a one-day term", () => {
    expect(
      SerialDate.monthRanges({
        startDate: ymd(2024, 1, 5),
        endDate: ymd(2024, 1, 5),
      }),
    ).toEqual([
      {
        month: 1,
        year: 2024,
        startDate: ymd(2024, 1, 5),
        endDate: ymd(2024, 1, 5),
      },
    ]);
  });

  it("starts the first entry on the term's start and ends the last on its end", () => {
    expect(
      SerialDate.monthRanges({
        startDate: ymd(2024, 1, 20),
        endDate: ymd(2024, 2, 10),
      }),
    ).toEqual([
      {
        month: 1,
        year: 2024,
        startDate: ymd(2024, 1, 20),
        endDate: ymd(2024, 1, 31),
      },
      {
        month: 2,
        year: 2024,
        startDate: ymd(2024, 2, 1),
        endDate: ymd(2024, 2, 10),
      },
    ]);
  });

  it("makes every interior entry a whole month, prorating to exactly 1", () => {
    const interiorMonths = SerialDate.monthRanges({
      startDate: ymd(2024, 1, 20),
      endDate: ymd(2024, 3, 10),
    }).slice(1, -1);

    expect(interiorMonths).toEqual([
      {
        month: 2,
        year: 2024,
        startDate: ymd(2024, 2, 1),
        endDate: ymd(2024, 2, 29),
      },
    ]);
    expect(
      interiorMonths.map((range) =>
        SerialDate.proratedMonthlyProportion(range),
      ),
    ).toEqual([1]);
  });

  it("walks December into January across a year boundary", () => {
    expect(
      SerialDate.monthRanges({
        startDate: ymd(2024, 12, 15),
        endDate: ymd(2025, 1, 15),
      }),
    ).toEqual([
      {
        month: 12,
        year: 2024,
        startDate: ymd(2024, 12, 15),
        endDate: ymd(2024, 12, 31),
      },
      {
        month: 1,
        year: 2025,
        startDate: ymd(2025, 1, 1),
        endDate: ymd(2025, 1, 15),
      },
    ]);
  });

  it("produces the months monthYearsOnAndBetween produces for the same span", () => {
    const term = { startDate: ymd(2024, 11, 17), endDate: ymd(2025, 4, 3) };

    expect(
      SerialDate.monthRanges(term).map(({ month, year }) => ({ month, year })),
    ).toEqual(
      SerialDate.monthYearsOnAndBetween({
        startMonthYear: SerialDate.monthYear(term.startDate),
        endMonthYear: SerialDate.monthYear(term.endDate),
      }),
    );
  });

  it("throws when the end is before the start", () => {
    expect(() =>
      SerialDate.monthRanges({
        startDate: ymd(2024, 3, 1),
        endDate: ymd(2024, 1, 1),
      }),
    ).toThrowError(/after end date/);
  });
});

describe("SerialDate.monthRanges through SerialDate.proratedMonthlyAmount", () => {
  it("splits 5 January to 15 March 2026 at $1,500 a month into three amounts", () => {
    const monthRanges = SerialDate.monthRanges({
      startDate: ymd(2026, 1, 5),
      endDate: ymd(2026, 3, 15),
    });

    expect(
      monthRanges.map(({ month, year, ...range }) => ({
        month,
        year,
        amount: SerialDate.proratedMonthlyAmount(1500, range),
      })),
    ).toEqual([
      { month: 1, year: 2026, amount: 1500 * (27 / 31) },
      { month: 2, year: 2026, amount: 1500 },
      { month: 3, year: 2026, amount: 1500 * (15 / 31) },
    ]);
  });
});
