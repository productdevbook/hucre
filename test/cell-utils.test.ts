import { describe, expect, it } from "vitest";
import { colToLetter, cellRef, rangeRef, parseCellRef } from "../src/index";

describe("colToLetter", () => {
  it("should convert 0 to A", () => {
    expect(colToLetter(0)).toBe("A");
  });

  it("should convert 25 to Z", () => {
    expect(colToLetter(25)).toBe("Z");
  });

  it("should convert 26 to AA", () => {
    expect(colToLetter(26)).toBe("AA");
  });

  it("should convert 701 to ZZ", () => {
    expect(colToLetter(701)).toBe("ZZ");
  });

  it("should convert 702 to AAA", () => {
    expect(colToLetter(702)).toBe("AAA");
  });
});

describe("cellRef", () => {
  it("should convert (0,0) to A1", () => {
    expect(cellRef(0, 0)).toBe("A1");
  });

  it("should convert (9,2) to C10", () => {
    expect(cellRef(9, 2)).toBe("C10");
  });
});

describe("rangeRef", () => {
  it("should convert (0,0,9,3) to A1:D10", () => {
    expect(rangeRef(0, 0, 9, 3)).toBe("A1:D10");
  });
});

describe("parseCellRef", () => {
  it("should parse A1 to {row:0, col:0}", () => {
    expect(parseCellRef("A1")).toEqual({ row: 0, col: 0 });
  });

  it("should parse Z1 to {row:0, col:25}", () => {
    expect(parseCellRef("Z1")).toEqual({ row: 0, col: 25 });
  });

  it("should parse AA15 to {row:14, col:26}", () => {
    expect(parseCellRef("AA15")).toEqual({ row: 14, col: 26 });
  });
});

describe("round-trip", () => {
  it("should round-trip parseCellRef(cellRef(r,c))", () => {
    const cases = [
      [0, 0],
      [5, 3],
      [14, 26],
      [99, 701],
      [0, 702],
    ] as const;

    for (const [r, c] of cases) {
      const ref = cellRef(r, c);
      const parsed = parseCellRef(ref);
      expect(parsed.row).toBe(r);
      expect(parsed.col).toBe(c);
    }
  });
});
