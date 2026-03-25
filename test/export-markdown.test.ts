import { describe, expect, it } from "vitest";
import { toMarkdown } from "../src/export/markdown";
import type { Sheet } from "../src/_types";

/** Helper to create a minimal sheet */
function makeSheet(rows: Sheet["rows"]): Sheet {
  return {
    name: "Sheet1",
    rows,
  };
}

describe("toMarkdown", () => {
  it("basic markdown table with header", () => {
    const sheet = makeSheet([
      ["Name", "Age"],
      ["Alice", 30],
      ["Bob", 25],
    ]);
    const md = toMarkdown(sheet);
    const lines = md.split("\n");
    expect(lines.length).toBe(4); // header + separator + 2 data rows
    expect(lines[0]).toContain("Name");
    expect(lines[0]).toContain("Age");
    expect(lines[1]).toMatch(/^\|[\s\-:]+\|[\s\-:]+\|$/);
    expect(lines[2]).toContain("Alice");
    expect(lines[2]).toContain("30");
    expect(lines[3]).toContain("Bob");
    expect(lines[3]).toContain("25");
  });

  it("alignment: numbers right, strings left", () => {
    const sheet = makeSheet([
      ["Name", "Score"],
      ["Alice", 100],
      ["Bob", 200],
    ]);
    const md = toMarkdown(sheet);
    const lines = md.split("\n");
    const separator = lines[1];
    // Second column (numbers) should have right alignment marker (---:)
    const cols = separator.split("|").filter((s) => s.length > 0);
    expect(cols[1].trimEnd()).toMatch(/:$/); // right-aligned
    expect(cols[0].trimEnd()).not.toMatch(/:$/); // left-aligned (no colon at end)
  });

  it("custom alignment", () => {
    const sheet = makeSheet([
      ["A", "B", "C"],
      ["1", "2", "3"],
    ]);
    const md = toMarkdown(sheet, {
      alignment: ["center", "right", "left"],
    });
    const lines = md.split("\n");
    const separator = lines[1];
    const cols = separator.split("|").filter((s) => s.length > 0);
    // Center: starts and ends with :
    expect(cols[0]).toMatch(/^:/);
    expect(cols[0]).toMatch(/:$/);
    // Right: ends with :
    expect(cols[1].trimEnd()).toMatch(/:$/);
    expect(cols[1]).not.toMatch(/^:/);
    // Left: no colons at end
    expect(cols[2].trimEnd()).not.toMatch(/:$/);
  });

  it("pipe character escaped in values", () => {
    const sheet = makeSheet([["Header"], ["value | with | pipes"]]);
    const md = toMarkdown(sheet);
    expect(md).toContain("value \\| with \\| pipes");
    // The output should still be valid markdown (pipes in content are escaped)
    const lines = md.split("\n");
    expect(lines.length).toBe(3);
  });

  it("null cells produce empty content", () => {
    const sheet = makeSheet([
      ["A", "B"],
      [null, "hello"],
      ["world", null],
    ]);
    const md = toMarkdown(sheet);
    const lines = md.split("\n");
    // Row with null in first column
    expect(lines[2]).toMatch(/^\|\s+\|/); // starts with empty cell
    // Row with null in second column
    expect(lines[3]).toMatch(/\|\s+\|$/); // ends with empty cell
  });

  it("date formatting as ISO date string", () => {
    const d = new Date(Date.UTC(2024, 0, 15));
    const sheet = makeSheet([["Date"], [d]]);
    const md = toMarkdown(sheet);
    expect(md).toContain("2024-01-15");
  });

  it("long values truncated", () => {
    const longStr = "a".repeat(60);
    const sheet = makeSheet([["Data"], [longStr]]);
    const md = toMarkdown(sheet, { maxWidth: 20 });
    expect(md).toContain("aaaaaaaaaaaaaaaaa...");
    expect(md).not.toContain(longStr);
  });

  it("maxWidth option respected", () => {
    const sheet = makeSheet([["Header"], ["abcdefghij"]]);
    const md = toMarkdown(sheet, { maxWidth: 7 });
    expect(md).toContain("abcd...");
  });

  it("single column", () => {
    const sheet = makeSheet([["Name"], ["Alice"], ["Bob"]]);
    const md = toMarkdown(sheet);
    const lines = md.split("\n");
    expect(lines.length).toBe(4);
    // Each line should start and end with |
    for (const line of lines) {
      expect(line.startsWith("|")).toBe(true);
      expect(line.endsWith("|")).toBe(true);
    }
  });

  it("single row (just header)", () => {
    const sheet = makeSheet([["Name", "Age", "City"]]);
    const md = toMarkdown(sheet);
    const lines = md.split("\n");
    // Should have header + separator only (no data rows)
    expect(lines.length).toBe(2);
    expect(lines[0]).toContain("Name");
    expect(lines[1]).toMatch(/-/);
  });

  it("empty sheet returns empty string", () => {
    const sheet = makeSheet([]);
    const md = toMarkdown(sheet);
    expect(md).toBe("");
  });

  it("no header row option", () => {
    const sheet = makeSheet([
      ["Alice", 30],
      ["Bob", 25],
    ]);
    const md = toMarkdown(sheet, { headerRow: false });
    const lines = md.split("\n");
    // Should have: generated header + separator + 2 data rows
    expect(lines.length).toBe(4);
    // Generated header should have column numbers
    expect(lines[0]).toContain("1");
    expect(lines[0]).toContain("2");
    // Data should be in rows after separator
    expect(lines[2]).toContain("Alice");
    expect(lines[3]).toContain("Bob");
  });

  it("boolean values rendered as true/false", () => {
    const sheet = makeSheet([["Flag"], [true], [false]]);
    const md = toMarkdown(sheet);
    expect(md).toContain("true");
    expect(md).toContain("false");
  });

  it("number values rendered correctly", () => {
    const sheet = makeSheet([["Value"], [0], [-1], [3.14]]);
    const md = toMarkdown(sheet);
    expect(md).toContain("0");
    expect(md).toContain("-1");
    expect(md).toContain("3.14");
  });

  it("columns are padded to consistent width", () => {
    const sheet = makeSheet([
      ["Name", "Score"],
      ["A", 1],
      ["LongName", 100],
    ]);
    const md = toMarkdown(sheet);
    const lines = md.split("\n");
    // All lines should have the same length (consistent padding)
    const lengths = lines.map((l) => l.length);
    expect(new Set(lengths).size).toBe(1);
  });

  it("mixed types detect alignment correctly", () => {
    const sheet = makeSheet([
      ["Col1", "Col2"],
      ["text", 42],
      ["more", 99],
    ]);
    const md = toMarkdown(sheet);
    const lines = md.split("\n");
    const separator = lines[1];
    const cols = separator.split("|").filter((s) => s.length > 0);
    // Col1 (strings) should be left-aligned
    expect(cols[0].trimEnd()).not.toMatch(/:$/);
    // Col2 (numbers) should be right-aligned
    expect(cols[1].trimEnd()).toMatch(/:$/);
  });

  it("all null column defaults to left alignment", () => {
    const sheet = makeSheet([["Header"], [null], [null]]);
    const md = toMarkdown(sheet);
    const lines = md.split("\n");
    const separator = lines[1];
    // Should still produce valid markdown
    expect(separator).toMatch(/-/);
  });

  it("maxWidth of 3 or less still works", () => {
    const sheet = makeSheet([["Hi"], ["abcdef"]]);
    const md = toMarkdown(sheet, { maxWidth: 3 });
    // With maxWidth 3, "abcdef" should be truncated to "abc" (no room for "...")
    expect(md).toContain("abc");
  });

  it("ragged rows (different lengths) handled", () => {
    const sheet = makeSheet([["A", "B", "C"], ["1"], ["2", "3"]]);
    const md = toMarkdown(sheet);
    const lines = md.split("\n");
    // All lines should have the same number of | characters
    const pipeCounts = lines.map((l) => (l.match(/\|/g) || []).length);
    expect(new Set(pipeCounts).size).toBe(1);
  });
});
