import { describe, expect, it } from "vitest";

import { SerialDate } from "./SerialDate";
import { Tim } from "./Tim";

function iso(date: Date): string {
  return date.toISOString();
}

describe("Tim.sheetsEpochUtcMs / Tim.msPerDay", () => {
  it("are SerialDate's epoch and day length", () => {
    expect(Tim.sheetsEpochUtcMs).toBe(SerialDate.sheetsEpochUtcMs);
    expect(Tim.sheetsEpochUtcMs).toBe(Date.UTC(1899, 11, 30));
    expect(Tim.msPerDay).toBe(86400000);
  });
});

describe("Tim.wallClockParts", () => {
  it("is the instant's wall-clock fields in the given zone, zero-padded", () => {
    const parts = Tim.wallClockParts(
      new Date("2024-03-15T02:30:05Z"),
      "America/Chicago",
    );

    expect(parts).toMatchObject({
      year: "2024",
      month: "03",
      day: "14",
      hour: "21",
      minute: "30",
      second: "05",
    });
  });

  it("writes midnight as hour 00, not 24", () => {
    const parts = Tim.wallClockParts(
      new Date("2024-03-15T05:00:00Z"),
      "America/Chicago",
    );

    expect(parts).toMatchObject({ day: "15", hour: "00", minute: "00" });
  });
});

describe("Tim.nowTimestamp", () => {
  it("formats the current wall-clock time as YYYY-MM-DD HH:MM:SS", () => {
    expect(Tim.nowTimestamp("America/Chicago")).toMatch(
      /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/,
    );
  });
});

describe("Tim.getTzOffsetMinutes", () => {
  it("is the zone's UTC offset at the instant, positive east of UTC", () => {
    const winter = new Date("2024-01-15T12:00:00Z");

    expect(Tim.getTzOffsetMinutes(winter, "UTC")).toBe(0);
    expect(Tim.getTzOffsetMinutes(winter, "America/Chicago")).toBe(-360);
    expect(Tim.getTzOffsetMinutes(winter, "Asia/Kolkata")).toBe(330);
    expect(Tim.getTzOffsetMinutes(winter, "Pacific/Kiritimati")).toBe(840);
    expect(Tim.getTzOffsetMinutes(winter, "Pacific/Pago_Pago")).toBe(-660);
  });

  it("follows daylight saving", () => {
    expect(
      Tim.getTzOffsetMinutes(
        new Date("2024-03-10T07:59:59Z"),
        "America/Chicago",
      ),
    ).toBe(-360);
    expect(
      Tim.getTzOffsetMinutes(
        new Date("2024-03-10T08:00:00Z"),
        "America/Chicago",
      ),
    ).toBe(-300);
    expect(
      Tim.getTzOffsetMinutes(new Date("2024-07-01T12:00:00Z"), "Europe/London"),
    ).toBe(60);
  });
});

describe("Tim.serialToDateTime", () => {
  it("reads the fraction as wall-clock time in the given zone", () => {
    const noonJan1 = 45292.5;

    expect(iso(Tim.serialToDateTime(noonJan1, "UTC"))).toBe(
      "2024-01-01T12:00:00.000Z",
    );
    expect(iso(Tim.serialToDateTime(noonJan1, "America/Chicago"))).toBe(
      "2024-01-01T18:00:00.000Z",
    );
    expect(iso(Tim.serialToDateTime(noonJan1, "Asia/Tokyo"))).toBe(
      "2024-01-01T03:00:00.000Z",
    );
  });

  it("uses the daylight-saving offset of a summer date", () => {
    const noonJul1 = 45474.5;

    expect(iso(Tim.serialToDateTime(noonJul1, "America/Chicago"))).toBe(
      "2024-07-01T17:00:00.000Z",
    );
  });

  it("rounds to the nearest millisecond", () => {
    const justAfterNoon = 45292.5 + 0.4 / 86400000;

    expect(iso(Tim.serialToDateTime(justAfterNoon, "UTC"))).toBe(
      "2024-01-01T12:00:00.000Z",
    );
  });
});

describe("Tim.dateTimeToSerial", () => {
  it("is the instant's wall-clock time in the zone as a fractional serial", () => {
    expect(
      Tim.dateTimeToSerial(new Date("2024-01-01T18:00:00Z"), "America/Chicago"),
    ).toBe(45292.5);
    expect(
      Tim.dateTimeToSerial(new Date("2024-07-01T17:00:00Z"), "America/Chicago"),
    ).toBe(45474.5);
    expect(
      Tim.dateTimeToSerial(new Date("2024-01-01T03:00:00Z"), "Asia/Tokyo"),
    ).toBe(45292.5);
  });

  it("round-trips through serialToDateTime", () => {
    const serial = 45400.75;
    const zone = "America/Chicago";

    expect(Tim.dateTimeToSerial(Tim.serialToDateTime(serial, zone), zone)).toBe(
      serial,
    );
  });
});

describe("Tim.addDaysTz", () => {
  it("keeps the wall-clock time across a spring-forward day", () => {
    const noonMar9 = new Date("2024-03-09T18:00:00Z");

    expect(iso(Tim.addDaysTz(noonMar9, 1, "America/Chicago"))).toBe(
      "2024-03-10T17:00:00.000Z",
    );
  });

  it("keeps the wall-clock time across a fall-back day, going backwards", () => {
    const noonNov4 = new Date("2024-11-04T18:00:00Z");

    expect(iso(Tim.addDaysTz(noonNov4, -2, "America/Chicago"))).toBe(
      "2024-11-02T17:00:00.000Z",
    );
  });

  it("rolls the month and year", () => {
    expect(
      iso(Tim.addDaysTz(new Date("2024-12-31T03:00:00Z"), 1, "Asia/Tokyo")),
    ).toBe("2025-01-01T03:00:00.000Z");
  });
});

describe("Tim.addMonthsTz", () => {
  it("keeps the wall-clock time across a fall-back month", () => {
    const noonOct15 = new Date("2024-10-15T17:00:00Z");

    expect(iso(Tim.addMonthsTz(noonOct15, 1, "America/Chicago"))).toBe(
      "2024-11-15T18:00:00.000Z",
    );
  });

  it("overflows a day the target month lacks into the next month", () => {
    const noonJan31 = new Date("2024-01-31T18:00:00Z");

    expect(iso(Tim.addMonthsTz(noonJan31, 1, "America/Chicago"))).toBe(
      "2024-03-02T18:00:00.000Z",
    );
  });

  it("goes backwards across a year", () => {
    expect(
      iso(Tim.addMonthsTz(new Date("2024-02-15T12:00:00Z"), -3, "UTC")),
    ).toBe("2023-11-15T12:00:00.000Z");
  });
});
