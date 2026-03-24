import { describe, expect, it } from "vitest";
import {
  dateToSerial,
  formatDate,
  isDateFormat,
  parseDate,
  serialToDate,
  serialToTime,
  timeToSerial,
} from "../src/_date";

// ── Helpers ──────────────────────────────────────────────────────────

/** Create a UTC Date from components */
function utc(y: number, m: number, d: number, h = 0, min = 0, s = 0, ms = 0): Date {
  return new Date(Date.UTC(y, m - 1, d, h, min, s, ms));
}

/** Assert that two Dates have the same UTC millisecond timestamp */
function expectSameDate(actual: Date, expected: Date): void {
  expect(actual.getTime()).toBe(expected.getTime());
}

// ── serialToDate (1900 system) ──────────────────────────────────────

describe("serialToDate (1900 system)", () => {
  it("serial 1 → Jan 1, 1900", () => {
    expectSameDate(serialToDate(1), utc(1900, 1, 1));
  });

  it("serial 2 → Jan 2, 1900", () => {
    expectSameDate(serialToDate(2), utc(1900, 1, 2));
  });

  it("serial 59 → Feb 28, 1900", () => {
    expectSameDate(serialToDate(59), utc(1900, 2, 28));
  });

  it("serial 60 → Feb 28, 1900 (Lotus bug — phantom 'Feb 29' mapped to Feb 28)", () => {
    // Feb 29, 1900 doesn't exist (1900 is NOT a leap year).
    // JavaScript's Date auto-corrects Date.UTC(1900, 1, 29) to Mar 1.
    // We map serial 60 to Feb 28 — the last valid date before the phantom day.
    // Callers should be aware that serial 60 is an Excel/Lotus artifact.
    const d = serialToDate(60);
    expect(d.getUTCFullYear()).toBe(1900);
    expect(d.getUTCMonth()).toBe(1); // February (0-based)
    expect(d.getUTCDate()).toBe(28);
  });

  it("serial 61 → Mar 1, 1900", () => {
    expectSameDate(serialToDate(61), utc(1900, 3, 1));
  });

  it("serial 62 → Mar 2, 1900", () => {
    expectSameDate(serialToDate(62), utc(1900, 3, 2));
  });

  it("serial 0 → Dec 31, 1899 (Excel 'Jan 0, 1900' quirk — time-only)", () => {
    expectSameDate(serialToDate(0), utc(1899, 12, 31));
  });

  it("serial 44197 → Jan 1, 2021", () => {
    expectSameDate(serialToDate(44197), utc(2021, 1, 1));
  });

  it("serial 45658 → Jan 1, 2025", () => {
    expectSameDate(serialToDate(45658), utc(2025, 1, 1));
  });

  it("serial 44197.5 → Jan 1, 2021 12:00:00", () => {
    expectSameDate(serialToDate(44197.5), utc(2021, 1, 1, 12, 0, 0));
  });

  it("serial 44197.75 → Jan 1, 2021 18:00:00", () => {
    expectSameDate(serialToDate(44197.75), utc(2021, 1, 1, 18, 0, 0));
  });

  it("serial 44197.25 → Jan 1, 2021 06:00:00", () => {
    expectSameDate(serialToDate(44197.25), utc(2021, 1, 1, 6, 0, 0));
  });

  it("serial 1.5 → Jan 1, 1900 12:00:00", () => {
    expectSameDate(serialToDate(1.5), utc(1900, 1, 1, 12, 0, 0));
  });

  it("serial 366 → Dec 31, 1900", () => {
    expectSameDate(serialToDate(366), utc(1900, 12, 31));
  });

  it("serial 367 → Jan 1, 1901", () => {
    expectSameDate(serialToDate(367), utc(1901, 1, 1));
  });

  it("serial 36526 → Jan 1, 2000", () => {
    expectSameDate(serialToDate(36526), utc(2000, 1, 1));
  });

  it("serial with time fraction preserves time", () => {
    const d = serialToDate(44197 + 14 / 24 + 30 / 1440); // Jan 1, 2021 14:30:00
    expect(d.getUTCHours()).toBe(14);
    expect(d.getUTCMinutes()).toBe(30);
  });

  it("very large serial (year 9999)", () => {
    // Dec 31, 9999 = serial 2958465 in Excel
    const d = serialToDate(2958465);
    expect(d.getUTCFullYear()).toBe(9999);
    expect(d.getUTCMonth()).toBe(11); // December
    expect(d.getUTCDate()).toBe(31);
  });

  it("serial 44197.9999999 does NOT roll over to next day", () => {
    const d = serialToDate(44197.9999999);
    expect(d.getUTCFullYear()).toBe(2021);
    expect(d.getUTCMonth()).toBe(0);
    expect(d.getUTCDate()).toBe(1);
  });
});

// ── serialToDate (1904 system) ──────────────────────────────────────

describe("serialToDate (1904 system)", () => {
  it("serial 0 → Jan 1, 1904", () => {
    expectSameDate(serialToDate(0, true), utc(1904, 1, 1));
  });

  it("serial 1 → Jan 2, 1904", () => {
    expectSameDate(serialToDate(1, true), utc(1904, 1, 2));
  });

  it("serial 42735 → Jan 1, 2021", () => {
    // 1904 system: Jan 1, 2021
    // Difference between 1900 and 1904 epochs: 1462 days
    // But we also need to account for the Lotus bug (+1).
    // In 1900 system: Jan 1, 2021 = 44197
    // 1904 serial = 1900 serial - 1462
    // 44197 - 1462 = 42735
    expectSameDate(serialToDate(42735, true), utc(2021, 1, 1));
  });

  it("serial 0.5 → Jan 1, 1904 12:00:00", () => {
    expectSameDate(serialToDate(0.5, true), utc(1904, 1, 1, 12, 0, 0));
  });

  it("no Lotus bug in 1904 system (serial 60 is a normal date)", () => {
    // Serial 60 in 1904 = Mar 1, 1904
    expectSameDate(serialToDate(60, true), utc(1904, 3, 1));
  });
});

// ── dateToSerial (1900 system) ──────────────────────────────────────

describe("dateToSerial (1900 system)", () => {
  it("Jan 1, 1900 → 1", () => {
    expect(dateToSerial(utc(1900, 1, 1))).toBe(1);
  });

  it("Jan 2, 1900 → 2", () => {
    expect(dateToSerial(utc(1900, 1, 2))).toBe(2);
  });

  it("Feb 28, 1900 → 59", () => {
    expect(dateToSerial(utc(1900, 2, 28))).toBe(59);
  });

  it("Mar 1, 1900 → 61 (skips the phantom Feb 29)", () => {
    expect(dateToSerial(utc(1900, 3, 1))).toBe(61);
  });

  it("Mar 2, 1900 → 62", () => {
    expect(dateToSerial(utc(1900, 3, 2))).toBe(62);
  });

  it("Dec 31, 1900 → 366", () => {
    expect(dateToSerial(utc(1900, 12, 31))).toBe(366);
  });

  it("Jan 1, 1901 → 367", () => {
    expect(dateToSerial(utc(1901, 1, 1))).toBe(367);
  });

  it("Jan 1, 2000 → 36526", () => {
    expect(dateToSerial(utc(2000, 1, 1))).toBe(36526);
  });

  it("Jan 1, 2021 → 44197", () => {
    expect(dateToSerial(utc(2021, 1, 1))).toBe(44197);
  });

  it("Jan 1, 2025 → 45658", () => {
    expect(dateToSerial(utc(2025, 1, 1))).toBe(45658);
  });

  it("Jan 1, 2021 12:00:00 → 44197.5", () => {
    expect(dateToSerial(utc(2021, 1, 1, 12, 0, 0))).toBe(44197.5);
  });

  it("Jan 1, 2021 06:00:00 → 44197.25", () => {
    expect(dateToSerial(utc(2021, 1, 1, 6, 0, 0))).toBe(44197.25);
  });

  it("Jan 1, 2021 18:00:00 → 44197.75", () => {
    expect(dateToSerial(utc(2021, 1, 1, 18, 0, 0))).toBe(44197.75);
  });
});

// ── dateToSerial (1904 system) ──────────────────────────────────────

describe("dateToSerial (1904 system)", () => {
  it("Jan 1, 1904 → 0", () => {
    expect(dateToSerial(utc(1904, 1, 1), true)).toBe(0);
  });

  it("Jan 2, 1904 → 1", () => {
    expect(dateToSerial(utc(1904, 1, 2), true)).toBe(1);
  });

  it("Jan 1, 2021 → 42735", () => {
    expect(dateToSerial(utc(2021, 1, 1), true)).toBe(42735);
  });
});

// ── Round-trip tests ────────────────────────────────────────────────

describe("round-trip: dateToSerial → serialToDate", () => {
  const testDates = [
    utc(1900, 1, 1),
    utc(1900, 2, 28),
    utc(1900, 3, 1),
    utc(1950, 6, 15),
    utc(1999, 12, 31),
    utc(2000, 1, 1),
    utc(2000, 2, 29), // real leap year
    utc(2021, 1, 1),
    utc(2021, 7, 4, 14, 30, 0),
    utc(2025, 1, 1),
    utc(2025, 12, 31, 23, 59, 59),
  ];

  for (const date of testDates) {
    it(`round-trip ${date.toISOString()} (1900 system)`, () => {
      const serial = dateToSerial(date);
      const result = serialToDate(serial);
      // Compare to millisecond precision
      expect(Math.abs(result.getTime() - date.getTime())).toBeLessThan(2);
    });
  }

  for (const date of testDates) {
    it(`round-trip ${date.toISOString()} (1904 system)`, () => {
      const serial = dateToSerial(date, true);
      const result = serialToDate(serial, true);
      expect(Math.abs(result.getTime() - date.getTime())).toBeLessThan(2);
    });
  }

  it("round-trip 100 random dates (1900 system)", () => {
    // Seed-independent: test a spread of dates from 1900 to 2100
    for (let i = 0; i < 100; i++) {
      const serial = 1 + Math.floor((i / 100) * 73050); // ~1900 to ~2100
      // Skip the Lotus bug serial
      if (serial === 60) continue;
      const date = serialToDate(serial);
      const backSerial = dateToSerial(date);
      expect(backSerial).toBe(serial);
    }
  });

  it("round-trip 100 random dates (1904 system)", () => {
    for (let i = 0; i < 100; i++) {
      const serial = Math.floor((i / 100) * 73050);
      const date = serialToDate(serial, true);
      const backSerial = dateToSerial(date, true);
      expect(backSerial).toBe(serial);
    }
  });
});

// ── isDateFormat ────────────────────────────────────────────────────

describe("isDateFormat", () => {
  describe("should return true for date/time formats", () => {
    const dateFmts = [
      "yyyy-mm-dd",
      "mm/dd/yyyy",
      "d-mmm-yy",
      "m/d/yy h:mm",
      "h:mm:ss",
      "h:mm AM/PM",
      "hh:mm:ss",
      "hh:mm",
      "yyyy",
      "dd/mm/yyyy",
      "dddd, mmmm d, yyyy",
      "d-mmm",
      "mmm-yy",
      "mmmm yyyy",
      "mm:ss", // minutes:seconds is a time format
      "[h]:mm:ss", // elapsed hours
      "yyyy/mm/dd hh:mm:ss",
      "d/m/yy",
      "yy-mm-dd",
    ];

    for (const fmt of dateFmts) {
      it(`"${fmt}" → true`, () => {
        expect(isDateFormat(fmt)).toBe(true);
      });
    }
  });

  describe("should return true for formats with locale prefixes", () => {
    it('"[$-409]m/d/yy h:mm AM/PM" → true', () => {
      expect(isDateFormat("[$-409]m/d/yy h:mm AM/PM")).toBe(true);
    });

    it('"[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy" → true', () => {
      expect(isDateFormat("[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy")).toBe(true);
    });
  });

  describe("should return true for built-in format IDs", () => {
    const dateIds = [14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47];
    for (const id of dateIds) {
      it(`format ID ${id} → true`, () => {
        expect(isDateFormat(String(id))).toBe(true);
      });
    }
  });

  describe("should return false for number formats", () => {
    const numFmts = [
      "General",
      "0",
      "0.00",
      "#,##0",
      "#,##0.00",
      "0%",
      "0.00%",
      "$#,##0",
      "$#,##0.00",
      "0.00E+00",
      "##0.0E+0",
      "@", // text
      "#,##0;(#,##0)", // number with negative
      "#,##0.00;[Red](#,##0.00)", // number with color
      "0.0",
    ];

    for (const fmt of numFmts) {
      it(`"${fmt}" → false`, () => {
        expect(isDateFormat(fmt)).toBe(false);
      });
    }
  });

  describe("should return false for empty/invalid inputs", () => {
    it('"" → false', () => {
      expect(isDateFormat("")).toBe(false);
    });
  });

  describe("edge cases for m/mm ambiguity", () => {
    it('"m" alone (no d/y context) → false', () => {
      // Standalone "m" without day/year is treated as a number format
      expect(isDateFormat("m")).toBe(false);
    });

    it('"mm" alone → false', () => {
      expect(isDateFormat("mm")).toBe(false);
    });

    it('"mmm" (month abbreviation) → true', () => {
      expect(isDateFormat("mmm")).toBe(true);
    });

    it('"mmmm" (month full name) → true', () => {
      expect(isDateFormat("mmmm")).toBe(true);
    });
  });
});

// ── formatDate ──────────────────────────────────────────────────────

describe("formatDate", () => {
  // Jan 15, 2021 14:30:00 UTC (a Friday)
  const date = utc(2021, 1, 15, 14, 30, 0);

  it('"yyyy-mm-dd" → "2021-01-15"', () => {
    expect(formatDate(date, "yyyy-mm-dd")).toBe("2021-01-15");
  });

  it('"mm/dd/yyyy" → "01/15/2021"', () => {
    expect(formatDate(date, "mm/dd/yyyy")).toBe("01/15/2021");
  });

  it('"d-mmm-yy" → "15-Jan-21"', () => {
    expect(formatDate(date, "d-mmm-yy")).toBe("15-Jan-21");
  });

  it('"dddd, mmmm d, yyyy" → "Friday, January 15, 2021"', () => {
    expect(formatDate(date, "dddd, mmmm d, yyyy")).toBe("Friday, January 15, 2021");
  });

  it('"h:mm:ss" → "14:30:00"', () => {
    expect(formatDate(date, "h:mm:ss")).toBe("14:30:00");
  });

  it('"h:mm AM/PM" → "2:30 PM"', () => {
    expect(formatDate(date, "h:mm AM/PM")).toBe("2:30 PM");
  });

  it('"hh:mm:ss" → "14:30:00"', () => {
    expect(formatDate(date, "hh:mm:ss")).toBe("14:30:00");
  });

  it("AM/PM with morning hour", () => {
    const morning = utc(2021, 1, 15, 9, 5, 0);
    expect(formatDate(morning, "h:mm AM/PM")).toBe("9:05 AM");
  });

  it("AM/PM with midnight (12 AM)", () => {
    const midnight = utc(2021, 1, 15, 0, 0, 0);
    expect(formatDate(midnight, "h:mm AM/PM")).toBe("12:00 AM");
  });

  it("AM/PM with noon (12 PM)", () => {
    const noon = utc(2021, 1, 15, 12, 0, 0);
    expect(formatDate(noon, "h:mm AM/PM")).toBe("12:00 PM");
  });

  it('"yy" → "21"', () => {
    expect(formatDate(date, "yy")).toBe("21");
  });

  it('"m/d/yy" → "1/15/21"', () => {
    expect(formatDate(date, "m/d/yy")).toBe("1/15/21");
  });

  it('"ddd" → "Fri"', () => {
    expect(formatDate(date, "ddd")).toBe("Fri");
  });

  it("formats with locale prefix are handled (prefix stripped)", () => {
    expect(formatDate(date, "[$-409]mm/dd/yyyy")).toBe("01/15/2021");
  });

  it('"m/d/yy h:mm" correctly uses m as month before d, and mm as minutes after h', () => {
    expect(formatDate(date, "m/d/yy h:mm")).toBe("1/15/21 14:30");
  });
});

// ── parseDate ───────────────────────────────────────────────────────

describe("parseDate", () => {
  it("ISO 8601 date only: '2021-01-15'", () => {
    const d = parseDate("2021-01-15");
    expect(d).not.toBeNull();
    expectSameDate(d!, utc(2021, 1, 15));
  });

  it("ISO 8601 with time and Z: '2021-01-15T14:30:00Z'", () => {
    const d = parseDate("2021-01-15T14:30:00Z");
    expect(d).not.toBeNull();
    expectSameDate(d!, utc(2021, 1, 15, 14, 30, 0));
  });

  it("ISO 8601 with timezone offset: '2021-01-15T14:30:00+05:30'", () => {
    const d = parseDate("2021-01-15T14:30:00+05:30");
    expect(d).not.toBeNull();
    // 14:30 IST = 09:00 UTC
    expectSameDate(d!, utc(2021, 1, 15, 9, 0, 0));
  });

  it("ISO 8601 with milliseconds: '2021-01-15T14:30:00.123Z'", () => {
    const d = parseDate("2021-01-15T14:30:00.123Z");
    expect(d).not.toBeNull();
    expectSameDate(d!, utc(2021, 1, 15, 14, 30, 0, 123));
  });

  it("US format: '01/15/2021'", () => {
    const d = parseDate("01/15/2021");
    expect(d).not.toBeNull();
    expectSameDate(d!, utc(2021, 1, 15));
  });

  it("US format without leading zeros: '1/5/2021'", () => {
    const d = parseDate("1/5/2021");
    expect(d).not.toBeNull();
    expectSameDate(d!, utc(2021, 1, 5));
  });

  it("invalid: 'not a date' → null", () => {
    expect(parseDate("not a date")).toBeNull();
  });

  it("invalid: '' → null", () => {
    expect(parseDate("")).toBeNull();
  });

  it("invalid: '   ' → null", () => {
    expect(parseDate("   ")).toBeNull();
  });

  it("trims whitespace", () => {
    const d = parseDate("  2021-01-15  ");
    expect(d).not.toBeNull();
    expectSameDate(d!, utc(2021, 1, 15));
  });
});

// ── serialToTime ────────────────────────────────────────────────────

describe("serialToTime", () => {
  it("0.0 → 00:00:00.000", () => {
    const t = serialToTime(0);
    expect(t.hours).toBe(0);
    expect(t.minutes).toBe(0);
    expect(t.seconds).toBe(0);
    expect(t.milliseconds).toBe(0);
  });

  it("0.5 → 12:00:00.000", () => {
    const t = serialToTime(0.5);
    expect(t.hours).toBe(12);
    expect(t.minutes).toBe(0);
    expect(t.seconds).toBe(0);
    expect(t.milliseconds).toBe(0);
  });

  it("0.75 → 18:00:00.000", () => {
    const t = serialToTime(0.75);
    expect(t.hours).toBe(18);
    expect(t.minutes).toBe(0);
    expect(t.seconds).toBe(0);
    expect(t.milliseconds).toBe(0);
  });

  it("0.25 → 06:00:00.000", () => {
    const t = serialToTime(0.25);
    expect(t.hours).toBe(6);
    expect(t.minutes).toBe(0);
    expect(t.seconds).toBe(0);
    expect(t.milliseconds).toBe(0);
  });

  it("23:59:59 (0.999988...) → 23:59:59", () => {
    // 23:59:59 = (23*3600 + 59*60 + 59) / 86400 = 86399/86400
    const serial = 86399 / 86400;
    const t = serialToTime(serial);
    expect(t.hours).toBe(23);
    expect(t.minutes).toBe(59);
    expect(t.seconds).toBe(59);
  });

  it("extracts time from serial with integer part (e.g. 44197.5)", () => {
    const t = serialToTime(44197.5);
    expect(t.hours).toBe(12);
    expect(t.minutes).toBe(0);
    expect(t.seconds).toBe(0);
  });

  it("0.6 → 14:24:00", () => {
    const t = serialToTime(0.6);
    expect(t.hours).toBe(14);
    expect(t.minutes).toBe(24);
    expect(t.seconds).toBe(0);
  });
});

// ── timeToSerial ────────────────────────────────────────────────────

describe("timeToSerial", () => {
  it("00:00:00 → 0", () => {
    expect(timeToSerial(0, 0, 0)).toBe(0);
  });

  it("12:00:00 → 0.5", () => {
    expect(timeToSerial(12, 0, 0)).toBe(0.5);
  });

  it("18:00:00 → 0.75", () => {
    expect(timeToSerial(18, 0, 0)).toBe(0.75);
  });

  it("06:00:00 → 0.25", () => {
    expect(timeToSerial(6, 0, 0)).toBe(0.25);
  });

  it("23:59:59 → ~0.999988", () => {
    const serial = timeToSerial(23, 59, 59);
    expect(serial).toBeCloseTo(86399 / 86400, 10);
  });

  it("defaults seconds and milliseconds to 0", () => {
    expect(timeToSerial(12, 0)).toBe(0.5);
  });

  it("includes milliseconds", () => {
    const serial = timeToSerial(0, 0, 0, 500);
    expect(serial).toBeCloseTo(500 / 86_400_000, 15);
  });
});

// ── serialToTime ↔ timeToSerial round-trip ──────────────────────────

describe("serialToTime ↔ timeToSerial round-trip", () => {
  const cases = [
    { h: 0, m: 0, s: 0 },
    { h: 6, m: 0, s: 0 },
    { h: 12, m: 0, s: 0 },
    { h: 18, m: 0, s: 0 },
    { h: 23, m: 59, s: 59 },
    { h: 14, m: 30, s: 45 },
    { h: 1, m: 1, s: 1 },
    { h: 8, m: 15, s: 30 },
  ];

  for (const { h, m, s } of cases) {
    it(`round-trip ${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}:${String(s).padStart(2, "0")}`, () => {
      const serial = timeToSerial(h, m, s);
      const result = serialToTime(serial);
      expect(result.hours).toBe(h);
      expect(result.minutes).toBe(m);
      expect(result.seconds).toBe(s);
    });
  }
});

// ── Timezone independence ───────────────────────────────────────────

describe("timezone independence", () => {
  it("serialToDate always returns UTC dates regardless of environment", () => {
    const d = serialToDate(44197);
    expect(d.getUTCFullYear()).toBe(2021);
    expect(d.getUTCMonth()).toBe(0);
    expect(d.getUTCDate()).toBe(1);
    expect(d.getUTCHours()).toBe(0);
    expect(d.getUTCMinutes()).toBe(0);
  });

  it("dateToSerial uses UTC components", () => {
    // Create a date in UTC and verify serial is based on UTC
    const d = new Date(Date.UTC(2021, 0, 1, 12, 0, 0));
    const serial = dateToSerial(d);
    expect(serial).toBe(44197.5);
  });
});

// ── Precision edge cases ────────────────────────────────────────────

describe("precision", () => {
  it("serial 44197.9999999 does not roll over to Jan 2", () => {
    const d = serialToDate(44197.9999999);
    expect(d.getUTCDate()).toBe(1);
    expect(d.getUTCMonth()).toBe(0);
    expect(d.getUTCFullYear()).toBe(2021);
  });

  it("dateToSerial produces clean values for midnight dates", () => {
    const serial = dateToSerial(utc(2021, 1, 1));
    expect(serial).toBe(44197);
    expect(serial % 1).toBe(0); // no fractional part
  });

  it("dateToSerial produces clean halves for noon", () => {
    const serial = dateToSerial(utc(2021, 1, 1, 12, 0, 0));
    expect(serial).toBe(44197.5);
  });
});
