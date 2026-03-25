import { describe, expect, it } from "vitest";
import { toJson } from "../src/export/json";
import type { Sheet } from "../src/_types";

/** Helper to create a minimal sheet */
function makeSheet(rows: Sheet["rows"], overrides?: Partial<Sheet>): Sheet {
  return {
    name: "Sheet1",
    rows,
    ...overrides,
  };
}

describe("toJson (#72)", () => {
  // ── objects format (default) ───────────────────────────────────

  it("should export as array of objects by default", () => {
    const sheet = makeSheet([
      ["Name", "Price"],
      ["Widget", 9.99],
      ["Gadget", 24.5],
    ]);
    const result = JSON.parse(toJson(sheet));
    expect(result).toEqual([
      { Name: "Widget", Price: 9.99 },
      { Name: "Gadget", Price: 24.5 },
    ]);
  });

  it("should handle empty sheet in objects format", () => {
    const sheet = makeSheet([]);
    const result = JSON.parse(toJson(sheet));
    expect(result).toEqual([]);
  });

  it("should handle header-only sheet", () => {
    const sheet = makeSheet([["A", "B", "C"]]);
    const result = JSON.parse(toJson(sheet));
    expect(result).toEqual([]);
  });

  it("should handle null values in objects format", () => {
    const sheet = makeSheet([
      ["Name", "Value"],
      ["Test", null],
    ]);
    const result = JSON.parse(toJson(sheet));
    expect(result).toEqual([{ Name: "Test", Value: null }]);
  });

  it("should handle rows shorter than header", () => {
    const sheet = makeSheet([["A", "B", "C"], ["x"]]);
    const result = JSON.parse(toJson(sheet));
    expect(result).toEqual([{ A: "x", B: null, C: null }]);
  });

  // ── arrays format ─────────────────────────────────────────────

  it("should export as arrays format", () => {
    const sheet = makeSheet([
      ["Name", "Price"],
      ["Widget", 9.99],
      ["Gadget", 24.5],
    ]);
    const result = JSON.parse(toJson(sheet, { format: "arrays" }));
    expect(result).toEqual({
      headers: ["Name", "Price"],
      data: [
        ["Widget", 9.99],
        ["Gadget", 24.5],
      ],
    });
  });

  it("should handle empty sheet in arrays format", () => {
    const sheet = makeSheet([]);
    const result = JSON.parse(toJson(sheet, { format: "arrays" }));
    expect(result).toEqual({ headers: [], data: [] });
  });

  // ── columns format ────────────────────────────────────────────

  it("should export as columns format", () => {
    const sheet = makeSheet([
      ["Name", "Price"],
      ["Widget", 9.99],
      ["Gadget", 24.5],
    ]);
    const result = JSON.parse(toJson(sheet, { format: "columns" }));
    expect(result).toEqual({
      Name: ["Widget", "Gadget"],
      Price: [9.99, 24.5],
    });
  });

  it("should handle empty sheet in columns format", () => {
    const sheet = makeSheet([]);
    const result = JSON.parse(toJson(sheet, { format: "columns" }));
    expect(result).toEqual({});
  });

  it("should handle header-only in columns format", () => {
    const sheet = makeSheet([["X", "Y"]]);
    const result = JSON.parse(toJson(sheet, { format: "columns" }));
    expect(result).toEqual({ X: [], Y: [] });
  });

  // ── headerRow option ──────────────────────────────────────────

  it("should support custom headerRow index", () => {
    const sheet = makeSheet([
      ["metadata", "row"],
      ["Name", "Price"],
      ["Widget", 9.99],
    ]);
    const result = JSON.parse(toJson(sheet, { headerRow: 1 }));
    expect(result).toEqual([{ Name: "Widget", Price: 9.99 }]);
  });

  // ── pretty option ─────────────────────────────────────────────

  it("should pretty print when option is true", () => {
    const sheet = makeSheet([["A"], [1]]);
    const result = toJson(sheet, { pretty: true });
    expect(result).toContain("\n");
    expect(result).toContain("  ");
  });

  it("should not pretty print by default", () => {
    const sheet = makeSheet([["A"], [1]]);
    const result = toJson(sheet);
    expect(result).not.toContain("\n");
  });

  // ── Date handling ─────────────────────────────────────────────

  it("should serialize Date values as ISO strings", () => {
    const date = new Date("2024-06-15T00:00:00.000Z");
    const sheet = makeSheet([["Date"], [date]]);
    const result = JSON.parse(toJson(sheet));
    expect(result[0].Date).toBe("2024-06-15T00:00:00.000Z");
  });

  // ── Boolean and mixed types ───────────────────────────────────

  it("should handle mixed types", () => {
    const sheet = makeSheet([
      ["str", "num", "bool", "nil"],
      ["hello", 42, true, null],
    ]);
    const result = JSON.parse(toJson(sheet));
    expect(result).toEqual([{ str: "hello", num: 42, bool: true, nil: null }]);
  });
});
