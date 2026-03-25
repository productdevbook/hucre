import { describe, it, expect } from "vitest";
import { formatValue } from "../src/_format";

// ── Default (no locale) ──────────────────────────────────────────────

describe("formatValue — default (no locale)", () => {
  it("uses dot as decimal separator", () => {
    expect(formatValue(1234.56, "0.00")).toBe("1234.56");
  });

  it("uses comma as thousands separator", () => {
    expect(formatValue(1234567, "#,##0")).toBe("1,234,567");
  });

  it("unchanged when no locale option provided", () => {
    expect(formatValue(1234.5, "#,##0.00")).toBe("1,234.50");
  });
});

// ── de-DE ────────────────────────────────────────────────────────────

describe("formatValue — de-DE locale", () => {
  const locale = "de-DE";

  it("uses comma as decimal separator", () => {
    expect(formatValue(1234.56, "0.00", { locale })).toBe("1234,56");
  });

  it("uses dot as thousands separator", () => {
    expect(formatValue(1234567, "#,##0", { locale })).toBe("1.234.567");
  });

  it("uses both locale-specific separators", () => {
    expect(formatValue(1234.5, "#,##0.00", { locale })).toBe("1.234,50");
  });

  it("handles zero correctly", () => {
    expect(formatValue(0, "#,##0.00", { locale })).toBe("0,00");
  });

  it("handles percentage", () => {
    expect(formatValue(0.15, "0.00%", { locale })).toBe("15,00%");
  });

  it("handles negative numbers", () => {
    expect(formatValue(-1234.5, "#,##0.00", { locale })).toBe("-1.234,50");
  });
});

// ── fr-FR ────────────────────────────────────────────────────────────

describe("formatValue — fr-FR locale", () => {
  const locale = "fr-FR";

  it("uses comma as decimal separator", () => {
    expect(formatValue(3.14, "0.00", { locale })).toBe("3,14");
  });

  it("uses non-breaking space as thousands separator", () => {
    const result = formatValue(1234567, "#,##0", { locale });
    // fr-FR uses \u00A0 (non-breaking space) as thousands separator
    expect(result).toBe("1\u00A0234\u00A0567");
  });

  it("combines both separators", () => {
    const result = formatValue(9876.54, "#,##0.00", { locale });
    expect(result).toBe("9\u00A0876,54");
  });
});

// ── tr-TR ────────────────────────────────────────────────────────────

describe("formatValue — tr-TR locale", () => {
  const locale = "tr-TR";

  it("uses comma as decimal separator", () => {
    expect(formatValue(42.5, "0.00", { locale })).toBe("42,50");
  });

  it("uses dot as thousands separator", () => {
    expect(formatValue(1000000, "#,##0", { locale })).toBe("1.000.000");
  });

  it("combines both separators", () => {
    expect(formatValue(12345.67, "#,##0.00", { locale })).toBe("12.345,67");
  });

  it("handles small numbers without thousands separator", () => {
    expect(formatValue(999, "#,##0.00", { locale })).toBe("999,00");
  });
});

// ── en-US explicit ───────────────────────────────────────────────────

describe("formatValue — en-US locale (explicit)", () => {
  const locale = "en-US";

  it("matches default behavior", () => {
    expect(formatValue(1234.5, "#,##0.00", { locale })).toBe("1,234.50");
  });

  it("decimal separator is dot", () => {
    expect(formatValue(3.14, "0.00", { locale })).toBe("3.14");
  });
});

// ── Unknown locale falls back to default ──────────────────────────────

describe("formatValue — unknown locale", () => {
  it("falls back to default (dot decimal, comma thousands)", () => {
    expect(formatValue(1234.5, "#,##0.00", { locale: "ja-JP" })).toBe("1,234.50");
  });
});

// ── Edge cases ────────────────────────────────────────────────────────

describe("formatValue — locale edge cases", () => {
  it("integer format with locale", () => {
    expect(formatValue(42, "0", { locale: "de-DE" })).toBe("42");
  });

  it("currency format with locale", () => {
    expect(formatValue(1234.5, "$#,##0.00", { locale: "de-DE" })).toBe("$1.234,50");
  });

  it("no format string still works with locale", () => {
    expect(formatValue(42, "General", { locale: "de-DE" })).toBe("42");
  });

  it("scientific notation with locale decimal", () => {
    const result = formatValue(1234, "0.00E+00", { locale: "de-DE" });
    expect(result).toBe("1,23E+03");
  });
});
