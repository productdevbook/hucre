import { describe, expect, it } from "vitest";
import {
  letterToCol,
  parseRange,
  isInRange,
  sheetToObjects,
  sheetToArrays,
  findCells,
  replaceCells,
  formatCsvValue,
  parseCsv,
  writeCsv,
} from "../src/index";
import type { Sheet, CellValue } from "../src/_types";

// ── Helpers ──────────────────────────────────────────────────────────

function makeSheet(overrides: Partial<Sheet> = {}): Sheet {
  return {
    name: "Sheet1",
    rows: [],
    ...overrides,
  };
}

// ── letterToCol ─────────────────────────────────────────────────────

describe("letterToCol", () => {
  it("should convert A to 0", () => {
    expect(letterToCol("A")).toBe(0);
  });

  it("should convert Z to 25", () => {
    expect(letterToCol("Z")).toBe(25);
  });

  it("should convert AA to 26", () => {
    expect(letterToCol("AA")).toBe(26);
  });

  it("should convert ZZ to 701", () => {
    expect(letterToCol("ZZ")).toBe(701);
  });

  it("should handle lowercase letters", () => {
    expect(letterToCol("a")).toBe(0);
    expect(letterToCol("z")).toBe(25);
    expect(letterToCol("aa")).toBe(26);
  });
});

// ── parseRange ──────────────────────────────────────────────────────

describe("parseRange", () => {
  it('should parse "A1:D10" correctly', () => {
    const result = parseRange("A1:D10");
    expect(result).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 9,
      endCol: 3,
    });
  });

  it('should parse "B2:E5" correctly', () => {
    const result = parseRange("B2:E5");
    expect(result).toEqual({
      startRow: 1,
      startCol: 1,
      endRow: 4,
      endCol: 4,
    });
  });

  it("should handle single-cell range", () => {
    const result = parseRange("C3");
    expect(result).toEqual({
      startRow: 2,
      startCol: 2,
      endRow: 2,
      endCol: 2,
    });
  });

  it('should parse "AA1:AB2" with multi-letter columns', () => {
    const result = parseRange("AA1:AB2");
    expect(result).toEqual({
      startRow: 0,
      startCol: 26,
      endRow: 1,
      endCol: 27,
    });
  });
});

// ── isInRange ───────────────────────────────────────────────────────

describe("isInRange", () => {
  const range = { startRow: 1, startCol: 1, endRow: 5, endCol: 5 };

  it("should return true for a cell inside the range", () => {
    expect(isInRange(3, 3, range)).toBe(true);
  });

  it("should return true for a cell on the top-left edge", () => {
    expect(isInRange(1, 1, range)).toBe(true);
  });

  it("should return true for a cell on the bottom-right edge", () => {
    expect(isInRange(5, 5, range)).toBe(true);
  });

  it("should return false for a cell above the range", () => {
    expect(isInRange(0, 3, range)).toBe(false);
  });

  it("should return false for a cell below the range", () => {
    expect(isInRange(6, 3, range)).toBe(false);
  });

  it("should return false for a cell to the left of the range", () => {
    expect(isInRange(3, 0, range)).toBe(false);
  });

  it("should return false for a cell to the right of the range", () => {
    expect(isInRange(3, 6, range)).toBe(false);
  });
});

// ── sheetToObjects ──────────────────────────────────────────────────

describe("sheetToObjects", () => {
  it("should convert rows to objects using first row as headers", () => {
    const sheet = makeSheet({
      rows: [
        ["Name", "Age", "City"],
        ["Alice", 30, "London"],
        ["Bob", 25, "Paris"],
      ],
    });
    const result = sheetToObjects(sheet);
    expect(result).toEqual([
      { Name: "Alice", Age: 30, City: "London" },
      { Name: "Bob", Age: 25, City: "Paris" },
    ]);
  });

  it("should use custom headerRow", () => {
    const sheet = makeSheet({
      rows: [["metadata row"], ["Name", "Score"], ["Alice", 100], ["Bob", 85]],
    });
    const result = sheetToObjects(sheet, { headerRow: 1 });
    expect(result).toEqual([
      { Name: "Alice", Score: 100 },
      { Name: "Bob", Score: 85 },
    ]);
  });

  it("should return empty array for empty sheet", () => {
    const sheet = makeSheet({ rows: [] });
    expect(sheetToObjects(sheet)).toEqual([]);
  });

  it("should handle null header values as empty string keys", () => {
    const sheet = makeSheet({
      rows: [
        ["Name", null, "Age"],
        ["Alice", "x", 30],
      ],
    });
    const result = sheetToObjects(sheet);
    expect(result).toEqual([{ Name: "Alice", "": "x", Age: 30 }]);
  });

  it("should fill missing columns with null", () => {
    const sheet = makeSheet({
      rows: [["A", "B", "C"], ["x"]],
    });
    const result = sheetToObjects(sheet);
    expect(result).toEqual([{ A: "x", B: null, C: null }]);
  });
});

// ── sheetToArrays ───────────────────────────────────────────────────

describe("sheetToArrays", () => {
  it("should split headers and data", () => {
    const sheet = makeSheet({
      rows: [
        ["Name", "Age"],
        ["Alice", 30],
        ["Bob", 25],
      ],
    });
    const result = sheetToArrays(sheet);
    expect(result.headers).toEqual(["Name", "Age"]);
    expect(result.data).toEqual([
      ["Alice", 30],
      ["Bob", 25],
    ]);
  });

  it("should handle empty sheet", () => {
    const sheet = makeSheet({ rows: [] });
    const result = sheetToArrays(sheet);
    expect(result.headers).toEqual([]);
    expect(result.data).toEqual([]);
  });

  it("should handle sheet with only headers", () => {
    const sheet = makeSheet({
      rows: [["A", "B", "C"]],
    });
    const result = sheetToArrays(sheet);
    expect(result.headers).toEqual(["A", "B", "C"]);
    expect(result.data).toEqual([]);
  });
});

// ── findCells ───────────────────────────────────────────────────────

describe("findCells", () => {
  it("should find cells by exact value", () => {
    const sheet = makeSheet({
      rows: [
        ["a", "b", "c"],
        ["d", "b", "f"],
        ["g", "h", "b"],
      ],
    });
    const result = findCells(sheet, "b");
    expect(result).toEqual([
      { row: 0, col: 1, value: "b" },
      { row: 1, col: 1, value: "b" },
      { row: 2, col: 2, value: "b" },
    ]);
  });

  it("should find cells by predicate", () => {
    const sheet = makeSheet({
      rows: [
        [1, 2, 3],
        [4, 5, 6],
        [7, 8, 9],
      ],
    });
    const result = findCells(sheet, (value) => typeof value === "number" && value > 5);
    expect(result).toEqual([
      { row: 1, col: 2, value: 6 },
      { row: 2, col: 0, value: 7 },
      { row: 2, col: 1, value: 8 },
      { row: 2, col: 2, value: 9 },
    ]);
  });

  it("should return empty array when nothing matches", () => {
    const sheet = makeSheet({
      rows: [
        ["a", "b"],
        ["c", "d"],
      ],
    });
    const result = findCells(sheet, "z");
    expect(result).toEqual([]);
  });

  it("should find null values", () => {
    const sheet = makeSheet({
      rows: [
        ["a", null],
        [null, "b"],
      ],
    });
    const result = findCells(sheet, null);
    expect(result).toEqual([
      { row: 0, col: 1, value: null },
      { row: 1, col: 0, value: null },
    ]);
  });

  it("predicate receives row and col", () => {
    const sheet = makeSheet({
      rows: [
        ["a", "b"],
        ["c", "d"],
      ],
    });
    const result = findCells(sheet, (_value, row, col) => row === 1 && col === 0);
    expect(result).toEqual([{ row: 1, col: 0, value: "c" }]);
  });
});

// ── replaceCells ────────────────────────────────────────────────────

describe("replaceCells", () => {
  it("should replace exact string values", () => {
    const sheet = makeSheet({
      rows: [
        ["hello", "world"],
        ["hello", "foo"],
      ],
    });
    const count = replaceCells(sheet, "hello", "hi");
    expect(count).toBe(2);
    expect(sheet.rows).toEqual([
      ["hi", "world"],
      ["hi", "foo"],
    ]);
  });

  it("should replace with regex", () => {
    const sheet = makeSheet({
      rows: [
        ["foo-123", "bar-456"],
        ["foo-789", "baz"],
      ],
    });
    const count = replaceCells(sheet, /foo-(\d+)/g, "replaced");
    expect(count).toBe(2);
    expect(sheet.rows).toEqual([
      ["replaced", "bar-456"],
      ["replaced", "baz"],
    ]);
  });

  it("should use regex capture groups when replace is a string", () => {
    const sheet = makeSheet({
      rows: [["hello world", "foo bar"]],
    });
    const count = replaceCells(sheet, /(\w+) (\w+)/g, "$2 $1");
    expect(count).toBe(2);
    expect(sheet.rows).toEqual([["world hello", "bar foo"]]);
  });

  it("should return 0 when nothing matches", () => {
    const sheet = makeSheet({
      rows: [
        ["a", "b"],
        ["c", "d"],
      ],
    });
    const count = replaceCells(sheet, "z", "x");
    expect(count).toBe(0);
  });

  it("should replace numeric values", () => {
    const sheet = makeSheet({
      rows: [
        [1, 2, 3],
        [4, 2, 6],
      ],
    });
    const count = replaceCells(sheet, 2, 99);
    expect(count).toBe(2);
    expect(sheet.rows).toEqual([
      [1, 99, 3],
      [4, 99, 6],
    ]);
  });
});

// ── CSV escapeFormulae ──────────────────────────────────────────────

describe("CSV escapeFormulae", () => {
  it('should prefix "=SUM(A1)" with single quote', () => {
    const result = formatCsvValue("=SUM(A1)", { escapeFormulae: true });
    expect(result).toBe("'=SUM(A1)");
  });

  it('should prefix "+cmd" with single quote', () => {
    const result = formatCsvValue("+cmd", { escapeFormulae: true });
    expect(result).toBe("'+cmd");
  });

  it('should prefix "-1+1" with single quote', () => {
    const result = formatCsvValue("-1+1", { escapeFormulae: true });
    expect(result).toBe("'-1+1");
  });

  it('should prefix "@SUM" with single quote', () => {
    const result = formatCsvValue("@SUM", { escapeFormulae: true });
    expect(result).toBe("'@SUM");
  });

  it("should not modify normal strings", () => {
    const result = formatCsvValue("hello world", { escapeFormulae: true });
    expect(result).toBe("hello world");
  });

  it("should not modify strings when escapeFormulae is false", () => {
    const result = formatCsvValue("=SUM(A1)", { escapeFormulae: false });
    expect(result).toBe("=SUM(A1)");
  });

  it("should not modify strings by default (escapeFormulae not set)", () => {
    const result = formatCsvValue("=SUM(A1)");
    expect(result).toBe("=SUM(A1)");
  });

  it("should work with writeCsv", () => {
    const rows: CellValue[][] = [
      ["Name", "Formula"],
      ["Alice", "=1+2"],
      ["Bob", "normal"],
    ];
    const csv = writeCsv(rows, { escapeFormulae: true });
    expect(csv).toContain("'=1+2");
    expect(csv).toContain("normal");
    expect(csv).not.toContain("'normal");
  });
});

// ── CSV maxRows ─────────────────────────────────────────────────────

describe("CSV maxRows", () => {
  it("should parse only first 3 rows from a 10-row CSV", () => {
    const lines: string[] = [];
    for (let i = 1; i <= 10; i++) {
      lines.push(`a${i},b${i},c${i}`);
    }
    const input = lines.join("\n");
    const result = parseCsv(input, { maxRows: 3 });
    expect(result.length).toBe(3);
    expect(result[0]).toEqual(["a1", "b1", "c1"]);
    expect(result[2]).toEqual(["a3", "b3", "c3"]);
  });

  it("should return all rows when maxRows is larger than row count", () => {
    const input = "a,b\nc,d\ne,f";
    const result = parseCsv(input, { maxRows: 100 });
    expect(result.length).toBe(3);
  });

  it("should return empty array when maxRows is 0", () => {
    const input = "a,b\nc,d";
    const result = parseCsv(input, { maxRows: 0 });
    expect(result.length).toBe(0);
  });

  it("should return all rows when maxRows is not set", () => {
    const input = "a,b\nc,d\ne,f";
    const result = parseCsv(input);
    expect(result.length).toBe(3);
  });
});
