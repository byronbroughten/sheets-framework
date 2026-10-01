declare const serialDate: unique symbol;
// A whole-day Sheets serial. Branded so a rent or a count can't be a date.
export type SerialDate = number & { readonly [serialDate]: true };

export interface MonthYear {
  month: number; // 1-12
  year: number;
}

export interface Ymd extends MonthYear {
  day: number;
}

export interface DateRange {
  startDate: SerialDate;
  endDate: SerialDate;
}

export interface DateInRange extends DateRange {
  date: SerialDate;
}

export interface MonthYearRange {
  startMonthYear: MonthYear;
  endMonthYear: MonthYear;
}

export interface FirstAndLastOfMonth {
  firstOfMonth: SerialDate;
  lastOfMonth: SerialDate;
}

export interface MonthRange extends MonthYear, DateRange {}

const monthAbbrevs = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
] as const;

export const SerialDate = {
  sheetsEpochUtcMs: Date.UTC(1899, 11, 30), // Dec 30, 1899, 00:00 UTC
  msPerDay: 86400000,
  isSerial(value: unknown): value is SerialDate {
    return typeof value === "number" && Number.isInteger(value);
  },
  validate(value: unknown): SerialDate {
    if (SerialDate.isSerial(value)) {
      return value;
    }
    throw new Error(`value "${String(value)}" is not a whole-day date serial`);
  },
  fromInstant(instant: Date, timeZone: string): SerialDate {
    const parts = new Intl.DateTimeFormat("en-US", {
      timeZone,
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    })
      .formatToParts(instant)
      .reduce<Record<string, string>>((acc, part) => {
        acc[part.type] = part.value;
        return acc;
      }, {});
    return SerialDate.fromYmd({
      year: Number(parts.year),
      month: Number(parts.month),
      day: Number(parts.day),
    });
  },
  fromYmd({ year, month, day }: Ymd): SerialDate {
    const utcMs = utcMsFromYmd({ year, month, day });
    const serial = (utcMs - SerialDate.sheetsEpochUtcMs) / SerialDate.msPerDay;
    if (!SerialDate.isSerial(serial)) {
      throw new Error(`${year}-${month}-${day} is not a real date.`);
    }
    return serial;
  },
  toYmd(date: SerialDate): Ymd {
    const utc = new Date(
      SerialDate.sheetsEpochUtcMs +
        SerialDate.validate(date) * SerialDate.msPerDay,
    );
    return {
      year: utc.getUTCFullYear(),
      month: utc.getUTCMonth() + 1,
      day: utc.getUTCDate(),
    };
  },
  toDayMonthYear(date: SerialDate): string {
    const { year, month, day } = SerialDate.toYmd(date);
    const monthAbbrev = monthAbbrevs[month - 1];
    if (monthAbbrev === undefined) {
      throw new Error(`month ${month} is not 1-12`);
    }
    return `${day} ${monthAbbrev} ${year}`;
  },
  addDays(date: SerialDate, days: number): SerialDate {
    return SerialDate.validate(SerialDate.validate(date) + days);
  },
  dayBefore(date: SerialDate): SerialDate {
    return SerialDate.addDays(date, -1);
  },
  addMonths(date: SerialDate, months: number): SerialDate {
    const { year, month, day } = SerialDate.toYmd(date);
    const monthCount = year * 12 + (month - 1) + months;
    const target = {
      month: (((monthCount % 12) + 12) % 12) + 1,
      year: Math.floor(monthCount / 12),
    };
    return SerialDate.fromYmd({
      ...target,
      day: Math.min(day, daysInMonthYear(target)),
    });
  },
  isSameOrAfter(date: SerialDate, referenceDate: SerialDate): boolean {
    return SerialDate.validate(date) >= SerialDate.validate(referenceDate);
  },
  isSameOrBefore(date: SerialDate, referenceDate: SerialDate): boolean {
    return SerialDate.validate(date) <= SerialDate.validate(referenceDate);
  },
  isOnOrBetween({ date, startDate, endDate }: DateInRange): boolean {
    if (SerialDate.validate(startDate) > SerialDate.validate(endDate)) {
      throw new Error("Start date cannot be after end date.");
    }
    return (
      SerialDate.isSameOrAfter(date, startDate) &&
      SerialDate.isSameOrBefore(date, endDate)
    );
  },
  monthYear(date: SerialDate): MonthYear {
    const { month, year } = SerialDate.toYmd(date);
    return { month, year };
  },
  isInMonthAndYear(date: SerialDate, { month, year }: MonthYear): boolean {
    const dateMonthYear = SerialDate.monthYear(date);
    return dateMonthYear.month === month && dateMonthYear.year === year;
  },
  monthYearsOnAndBetween({
    startMonthYear,
    endMonthYear,
  }: MonthYearRange): MonthYear[] {
    const monthYears: MonthYear[] = [];
    let current = startMonthYear;
    while (
      current.year < endMonthYear.year ||
      (current.year === endMonthYear.year &&
        current.month <= endMonthYear.month)
    ) {
      monthYears.push(current);
      current = nextMonthYear(current);
    }
    return monthYears;
  },
  firstDayOfMonth(date: SerialDate): SerialDate {
    return SerialDate.firstDayOfMonthYear(SerialDate.monthYear(date));
  },
  lastDayOfMonth(date: SerialDate): SerialDate {
    return SerialDate.lastDayOfMonthYear(SerialDate.monthYear(date));
  },
  firstAndLastDayOfMonth(date: SerialDate): FirstAndLastOfMonth {
    return SerialDate.firstAndLastDayOfMonthYear(SerialDate.monthYear(date));
  },
  firstDayOfNextMonth(date: SerialDate): SerialDate {
    return SerialDate.firstDayOfMonthYear(
      nextMonthYear(SerialDate.monthYear(date)),
    );
  },
  firstDayOfMonthYear({ month, year }: MonthYear): SerialDate {
    return SerialDate.fromYmd({ month, year, day: 1 });
  },
  lastDayOfMonthYear(monthYear: MonthYear): SerialDate {
    return SerialDate.dayBefore(
      SerialDate.firstDayOfMonthYear(nextMonthYear(monthYear)),
    );
  },
  firstAndLastDayOfMonthYear(monthYear: MonthYear): FirstAndLastOfMonth {
    return {
      firstOfMonth: SerialDate.firstDayOfMonthYear(monthYear),
      lastOfMonth: SerialDate.lastDayOfMonthYear(monthYear),
    };
  },
  monthRanges(term: DateRange): MonthRange[] {
    validateDateOrder(term);
    return SerialDate.monthYearsOnAndBetween({
      startMonthYear: SerialDate.monthYear(term.startDate),
      endMonthYear: SerialDate.monthYear(term.endDate),
    }).map((monthYear) => {
      const { firstOfMonth, lastOfMonth } =
        SerialDate.firstAndLastDayOfMonthYear(monthYear);
      return {
        ...monthYear,
        startDate: SerialDate.validate(Math.max(term.startDate, firstOfMonth)),
        endDate: SerialDate.validate(Math.min(term.endDate, lastOfMonth)),
      };
    });
  },
  proratedMonthlyProportion(range: DateRange): number {
    const monthYear = validateSingleMonth(range);
    return (range.endDate - range.startDate + 1) / daysInMonthYear(monthYear);
  },
  proratedMonthlyAmount(amount: number, range: DateRange): number {
    return SerialDate.proratedMonthlyProportion(range) * amount;
  },
};

// NaN when the calendar has no such day, so fromYmd throws instead of overflowing.
function utcMsFromYmd({ year, month, day }: Ymd): number {
  if (![year, month, day].every((part) => Number.isInteger(part))) {
    return NaN;
  }
  const utc = new Date(0);
  utc.setUTCFullYear(year, month - 1, day);
  utc.setUTCHours(0, 0, 0, 0);
  if (
    utc.getUTCFullYear() !== year ||
    utc.getUTCMonth() !== month - 1 ||
    utc.getUTCDate() !== day
  ) {
    return NaN;
  }
  return utc.getTime();
}

function daysInMonthYear(monthYear: MonthYear): number {
  const { firstOfMonth, lastOfMonth } =
    SerialDate.firstAndLastDayOfMonthYear(monthYear);
  return lastOfMonth - firstOfMonth + 1;
}

function nextMonthYear({ month, year }: MonthYear): MonthYear {
  if (month === 12) {
    return { month: 1, year: year + 1 };
  }
  return { month: month + 1, year };
}

function validateDateOrder({ startDate, endDate }: DateRange): void {
  if (SerialDate.validate(startDate) > SerialDate.validate(endDate)) {
    throw new Error("Start date cannot be after end date.");
  }
}

// Hands back the month it proved, so the caller doesn't derive it twice.
function validateSingleMonth(range: DateRange): MonthYear {
  validateDateOrder(range);
  const start = SerialDate.monthYear(range.startDate);
  const end = SerialDate.monthYear(range.endDate);
  if (start.month !== end.month || start.year !== end.year) {
    throw new Error(
      `A prorated range must lie in one month, but ${start.year}-${start.month} and ${end.year}-${end.month} differ.`,
    );
  }
  return start;
}
