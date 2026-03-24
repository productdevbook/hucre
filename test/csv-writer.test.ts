import { describe, expect, it } from "vitest";
import { writeCsv, writeCsvObjects, formatCsvValue } from "../src/csv/index";
import { parseCsv } from "../src/csv/index";
import type { CellValue } from "../src/_types";

describe("formatCsvValue", () => {
  it("should format null as empty string", () => {
    expect(formatCsvValue(null)).toBe("");
  });

  it("should format boolean true", () => {
    expect(formatCsvValue(true)).toBe("true");
  });

  it("should format boolean false", () => {
    expect(formatCsvValue(false)).toBe("false");
  });

  it("should format integers", () => {
    expect(formatCsvValue(42)).toBe("42");
  });

  it("should format floats", () => {
    expect(formatCsvValue(3.14)).toBe("3.14");
  });

  it("should format Date as ISO 8601", () => {
    const d = new Date("2024-01-15T10:30:00.000Z");
    expect(formatCsvValue(d)).toBe("2024-01-15T10:30:00.000Z");
  });

  it("should format strings without quoting when not needed", () => {
    expect(formatCsvValue("hello")).toBe("hello");
  });

  it("should quote strings containing comma", () => {
    expect(formatCsvValue("hello, world")).toBe('"hello, world"');
  });

  it("should escape quotes in strings", () => {
    expect(formatCsvValue('say "hi"')).toBe('"say ""hi"""');
  });
});

describe("writeCsv", () => {
  // ── Basic writing ──────────────────────────────────────────────

  it("should write simple array to CSV", () => {
    const result = writeCsv([
      ["a", "b", "c"],
      [1, 2, 3],
    ]);
    expect(result).toBe("a,b,c\n1,2,3");
  });

  it("should return empty string for empty array", () => {
    expect(writeCsv([])).toBe("");
  });

  it("should write single row", () => {
    expect(writeCsv([["a", "b", "c"]])).toBe("a,b,c");
  });

  // ── Quoting ────────────────────────────────────────────────────

  it("should quote field containing delimiter", () => {
    const result = writeCsv([["hello, world", "foo"]]);
    expect(result).toBe('"hello, world",foo');
  });

  it("should quote field containing quote character", () => {
    const result = writeCsv([['say "hi"', "foo"]]);
    expect(result).toBe('"say ""hi""",foo');
  });

  it("should quote field containing newline", () => {
    const result = writeCsv([["line1\nline2", "foo"]]);
    expect(result).toBe('"line1\nline2",foo');
  });

  it("should quote all fields when quoteStyle is 'all'", () => {
    const result = writeCsv([["a", "b"]], { quoteStyle: "all" });
    expect(result).toBe('"a","b"');
  });

  it("should not quote any fields when quoteStyle is 'none'", () => {
    const result = writeCsv([["hello, world", "foo"]], { quoteStyle: "none" });
    expect(result).toBe("hello, world,foo");
  });

  it("should quote only when required by default", () => {
    const result = writeCsv([["hello", "hi, there"]]);
    expect(result).toBe('hello,"hi, there"');
  });

  // ── Custom delimiter ───────────────────────────────────────────

  it("should use custom delimiter (semicolon)", () => {
    const result = writeCsv(
      [
        ["a", "b", "c"],
        [1, 2, 3],
      ],
      { delimiter: ";" },
    );
    expect(result).toBe("a;b;c\n1;2;3");
  });

  // ── Custom line separator ──────────────────────────────────────

  it("should use custom line separator (\\r\\n)", () => {
    const result = writeCsv(
      [
        ["a", "b"],
        [1, 2],
      ],
      { lineSeparator: "\r\n" },
    );
    expect(result).toBe("a,b\r\n1,2");
  });

  // ── BOM ────────────────────────────────────────────────────────

  it("should prepend UTF-8 BOM when bom option is true", () => {
    const result = writeCsv([["a", "b"]], { bom: true });
    expect(result.charCodeAt(0)).toBe(0xfeff);
    expect(result.slice(1)).toBe("a,b");
  });

  it("should not include BOM by default", () => {
    const result = writeCsv([["a", "b"]]);
    expect(result.charCodeAt(0)).not.toBe(0xfeff);
  });

  // ── Null values ────────────────────────────────────────────────

  it("should write null as empty string by default", () => {
    const result = writeCsv([["a", null, "c"]]);
    expect(result).toBe("a,,c");
  });

  it("should use custom null value", () => {
    const result = writeCsv([["a", null, "c"]], { nullValue: "NULL" });
    expect(result).toBe("a,NULL,c");
  });

  // ── Date formatting ────────────────────────────────────────────

  it("should format dates as ISO 8601 by default", () => {
    const d = new Date("2024-01-15T00:00:00.000Z");
    const result = writeCsv([[d]]);
    expect(result).toBe("2024-01-15T00:00:00.000Z");
  });

  it("should use custom date format", () => {
    const d = new Date(2024, 0, 15, 10, 30, 0); // local time
    const result = writeCsv([[d]], { dateFormat: "YYYY-MM-DD" });
    expect(result).toBe("2024-01-15");
  });

  // ── Boolean formatting ─────────────────────────────────────────

  it("should format booleans as true/false", () => {
    const result = writeCsv([[true, false]]);
    expect(result).toBe("true,false");
  });

  // ── Number formatting ──────────────────────────────────────────

  it("should not use scientific notation for large integers", () => {
    const result = writeCsv([[1e16]]);
    expect(result).not.toContain("e");
    expect(result).toBe("10000000000000000");
  });

  it("should preserve decimal precision", () => {
    const result = writeCsv([[3.14159265358979]]);
    expect(result).toBe("3.14159265358979");
  });

  // ── Headers ────────────────────────────────────────────────────

  it("should write explicit headers", () => {
    const result = writeCsv(
      [
        [1, 2],
        [3, 4],
      ],
      { headers: ["a", "b"] },
    );
    expect(result).toBe("a,b\n1,2\n3,4");
  });

  // ── Unicode content ────────────────────────────────────────────

  it("should handle unicode content", () => {
    const result = writeCsv([
      ["名前", "年齢"],
      ["太郎", 25],
    ]);
    expect(result).toBe("名前,年齢\n太郎,25");
  });

  it("should handle emoji content", () => {
    const result = writeCsv([["🎉", "🚀"]]);
    expect(result).toBe("🎉,🚀");
  });
});

describe("writeCsvObjects", () => {
  it("should write objects with auto headers from keys", () => {
    const result = writeCsvObjects([
      { name: "Alice", age: 30 },
      { name: "Bob", age: 25 },
    ]);
    expect(result).toBe("name,age\nAlice,30\nBob,25");
  });

  it("should write objects with explicit headers", () => {
    const result = writeCsvObjects(
      [
        { name: "Alice", age: 30 },
        { name: "Bob", age: 25 },
      ],
      { headers: ["age", "name"] },
    );
    expect(result).toBe("age,name\n30,Alice\n25,Bob");
  });

  it("should write objects with column ordering via headers", () => {
    const result = writeCsvObjects([{ z: 1, a: 2, m: 3 }], { headers: ["a", "m", "z"] });
    expect(result).toBe("a,m,z\n2,3,1");
  });

  it("should handle missing keys as null", () => {
    const result = writeCsvObjects([{ name: "Alice" } as Record<string, CellValue>], {
      headers: ["name", "age"],
    });
    expect(result).toBe("name,age\nAlice,");
  });

  it("should return empty string for empty data array", () => {
    expect(writeCsvObjects([])).toBe("");
  });

  it("should return only BOM for empty data with bom option", () => {
    const result = writeCsvObjects([], { bom: true });
    expect(result).toBe("\uFEFF");
  });

  it("should write objects without header row when headers is false", () => {
    const result = writeCsvObjects(
      [
        { name: "Alice", age: 30 },
        { name: "Bob", age: 25 },
      ],
      { headers: false },
    );
    expect(result).toBe("Alice,30\nBob,25");
  });
});

describe("round-trip", () => {
  it("should round-trip simple data", () => {
    const original: CellValue[][] = [
      ["name", "value"],
      ["hello", "world"],
      ["foo", "bar"],
    ];
    const csv = writeCsv(original);
    const parsed = parseCsv(csv);
    expect(parsed).toEqual(original);
  });

  it("should round-trip data with special characters", () => {
    const original: CellValue[][] = [["text"], ["hello, world"], ['say "hi"'], ["line1\nline2"]];
    const csv = writeCsv(original);
    const parsed = parseCsv(csv);
    expect(parsed).toEqual(original);
  });

  it("should round-trip with null values", () => {
    const original: CellValue[][] = [
      ["a", null, "c"],
      [null, "b", null],
    ];
    const csv = writeCsv(original);
    const parsed = parseCsv(csv);
    // nulls become empty strings after round-trip
    expect(parsed).toEqual([
      ["a", "", "c"],
      ["", "b", ""],
    ]);
  });

  it("should round-trip with BOM", () => {
    const original: CellValue[][] = [
      ["a", "b"],
      ["1", "2"],
    ];
    const csv = writeCsv(original, { bom: true });
    const parsed = parseCsv(csv); // BOM stripped by default
    expect(parsed).toEqual(original);
  });

  it("should round-trip with semicolon delimiter", () => {
    const original: CellValue[][] = [
      ["a", "b", "c"],
      ["1", "2", "3"],
    ];
    const csv = writeCsv(original, { delimiter: ";" });
    const parsed = parseCsv(csv, { delimiter: ";" });
    expect(parsed).toEqual(original);
  });

  it("should round-trip unicode content", () => {
    const original: CellValue[][] = [
      ["名前", "emoji"],
      ["太郎", "🎉"],
      ["أحمد", "🚀"],
    ];
    const csv = writeCsv(original);
    const parsed = parseCsv(csv);
    expect(parsed).toEqual(original);
  });
});
