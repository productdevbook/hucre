import { describe, it, expect } from "vitest";
import { formatValue } from "../src/_format";
import { dateToSerial } from "../src/_date";

// ── Helpers ──────────────────────────────────────────────────────────

/** Serial number for Jan 1, 2021 in the 1900 date system */
const JAN_1_2021 = dateToSerial(new Date(Date.UTC(2021, 0, 1)));

// ── General Format ──────────────────────────────────────────────────

describe("formatValue — General", () => {
  it("formats number as string", () => {
    expect(formatValue(42, "General")).toBe("42");
  });

  it("formats string as-is", () => {
    expect(formatValue("hello", "General")).toBe("hello");
  });

  it("handles empty/missing format as General", () => {
    expect(formatValue(42, "")).toBe("42");
  });
});

// ── Null / Undefined / Boolean ──────────────────────────────────────

describe("formatValue — special values", () => {
  it("null → empty string", () => {
    expect(formatValue(null, "0.00")).toBe("");
  });

  it("undefined → empty string", () => {
    expect(formatValue(undefined, "0.00")).toBe("");
  });

  it("boolean true → 'TRUE'", () => {
    expect(formatValue(true, "0.00")).toBe("TRUE");
  });

  it("boolean false → 'FALSE'", () => {
    expect(formatValue(false, "0")).toBe("FALSE");
  });
});

// ── Number Formats ──────────────────────────────────────────────────

describe("formatValue — Number formats", () => {
  it("0 format: rounds to integer", () => {
    expect(formatValue(42.567, "0")).toBe("43");
  });

  it("0.00 format: two decimal places", () => {
    expect(formatValue(42.567, "0.00")).toBe("42.57");
  });

  it("#,##0 format: thousand separators", () => {
    expect(formatValue(1234567, "#,##0")).toBe("1,234,567");
  });

  it("#,##0.00 format: thousand separators with decimals", () => {
    expect(formatValue(1234.5, "#,##0.00")).toBe("1,234.50");
  });

  it("0.00 with small number", () => {
    expect(formatValue(0.5, "0.00")).toBe("0.50");
  });

  it("#,##0 with zero", () => {
    expect(formatValue(0, "#,##0")).toBe("0");
  });
});

// ── Currency ────────────────────────────────────────────────────────

describe("formatValue — Currency", () => {
  it("$#,##0.00 format", () => {
    expect(formatValue(1234.5, "$#,##0.00")).toBe("$1,234.50");
  });

  it("currency with quoted symbol", () => {
    expect(formatValue(1234.5, '"€"#,##0.00')).toBe("€1,234.50");
  });
});

// ── Percentage ──────────────────────────────────────────────────────

describe("formatValue — Percentage", () => {
  it("0% format: multiply by 100", () => {
    expect(formatValue(0.15, "0%")).toBe("15%");
  });

  it("0.00% format: with decimals", () => {
    expect(formatValue(0.1567, "0.00%")).toBe("15.67%");
  });
});

// ── Scientific Notation ─────────────────────────────────────────────

describe("formatValue — Scientific", () => {
  it("0.00E+00 format", () => {
    expect(formatValue(1234, "0.00E+00")).toBe("1.23E+03");
  });

  it("0.00E+00 with small number", () => {
    expect(formatValue(0.005, "0.00E+00")).toBe("5.00E-03");
  });
});

// ── Text Format ─────────────────────────────────────────────────────

describe("formatValue — Text", () => {
  it("@ format returns value as string", () => {
    expect(formatValue(42, "@")).toBe("42");
  });

  it("@ with string value", () => {
    expect(formatValue("hello", "@")).toBe("hello");
  });
});

// ── Date Formats ────────────────────────────────────────────────────

describe("formatValue — Date formats", () => {
  it("yyyy-mm-dd with serial number", () => {
    const result = formatValue(JAN_1_2021, "yyyy-mm-dd");
    expect(result).toBe("2021-01-01");
  });

  it("m/d/yy with serial number", () => {
    const result = formatValue(JAN_1_2021, "m/d/yy");
    expect(result).toBe("1/1/21");
  });
});

// ── Multi-section Formats ───────────────────────────────────────────

describe("formatValue — Sections", () => {
  it("positive;negative: positive value uses first section", () => {
    expect(formatValue(42, "0.00;(0.00)")).toBe("42.00");
  });

  it("positive;negative: negative value uses second section (abs)", () => {
    expect(formatValue(-42, "0.00;(0.00)")).toBe("(42.00)");
  });

  it("positive;negative;zero: zero value uses third section", () => {
    expect(formatValue(0, '0.00;(0.00);"-"')).toBe("-");
  });

  it("four sections: string value uses fourth section", () => {
    expect(formatValue("hello", '0.00;(0.00);"-";@" text"')).toBe("hello text");
  });
});

// ── Color Codes ─────────────────────────────────────────────────────

describe("formatValue — Color codes", () => {
  it("[Red]0.00 strips color code", () => {
    expect(formatValue(42.567, "[Red]0.00")).toBe("42.57");
  });

  it("[Blue]#,##0 strips color code", () => {
    expect(formatValue(1234, "[Blue]#,##0")).toBe("1,234");
  });
});

// ── Locale Prefix ───────────────────────────────────────────────────

describe("formatValue — Locale prefix", () => {
  it("[$-409] is stripped", () => {
    expect(formatValue(1234.5, "[$-409]#,##0.00")).toBe("1,234.50");
  });

  it("locale with currency [$€-407]", () => {
    expect(formatValue(1234.5, "[$€-407]#,##0.00")).toBe("1,234.50");
  });
});

// ── Fractions ───────────────────────────────────────────────────────

describe("formatValue — Fractions", () => {
  it("# ?/? format for simple fraction", () => {
    const result = formatValue(1.5, "# ?/?");
    expect(result).toContain("1");
    expect(result).toContain("/");
    expect(result).toContain("2");
  });
});
