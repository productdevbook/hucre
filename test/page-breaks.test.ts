import { describe, it, expect } from "vitest";
import { writeXlsx } from "../src/xlsx/writer";
import { readXlsx } from "../src/xlsx/reader";
import { createStylesCollector } from "../src/xlsx/styles-writer";
import { createSharedStrings, writeWorksheetXml } from "../src/xlsx/worksheet-writer";
import { parseXml } from "../src/xml/parser";
import type { WriteSheet } from "../src/_types";

// ── Helpers ──────────────────────────────────────────────────────────

function writeXml(sheet: WriteSheet): string {
  const styles = createStylesCollector();
  const ss = createSharedStrings();
  const result = writeWorksheetXml(sheet, styles, ss);
  return result.xml;
}

function findChild(el: { children: Array<unknown> }, localName: string): any {
  return el.children.find((c: any) => typeof c !== "string" && (c.local || c.tag) === localName);
}

function findChildren(el: { children: Array<unknown> }, localName: string): any[] {
  return el.children.filter((c: any) => typeof c !== "string" && (c.local || c.tag) === localName);
}

// ── Row Breaks Writing ──────────────────────────────────────────────

describe("page breaks — row breaks writing", () => {
  it("writes row breaks with correct XML structure", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["A"], ["B"], ["C"]],
      rowBreaks: [9, 24], // 0-based: break after row 10 and 25
    };

    const xml = writeXml(sheet);
    const doc = parseXml(xml);

    const rb = findChild(doc, "rowBreaks");
    expect(rb).toBeDefined();
    expect(rb.attrs["count"]).toBe("2");
    expect(rb.attrs["manualBreakCount"]).toBe("2");

    const brks = findChildren(rb, "brk");
    expect(brks).toHaveLength(2);

    // id = 0-based + 1 = 1-based
    expect(brks[0].attrs["id"]).toBe("10");
    expect(brks[0].attrs["max"]).toBe("16383");
    expect(brks[0].attrs["man"]).toBe("1");

    expect(brks[1].attrs["id"]).toBe("25");
    expect(brks[1].attrs["max"]).toBe("16383");
    expect(brks[1].attrs["man"]).toBe("1");
  });

  it("sorts row breaks in output", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["A"]],
      rowBreaks: [24, 9], // unsorted
    };

    const xml = writeXml(sheet);
    const doc = parseXml(xml);

    const rb = findChild(doc, "rowBreaks");
    const brks = findChildren(rb, "brk");

    // Should be sorted: 10, 25 (1-based)
    expect(brks[0].attrs["id"]).toBe("10");
    expect(brks[1].attrs["id"]).toBe("25");
  });
});

// ── Column Breaks Writing ───────────────────────────────────────────

describe("page breaks — column breaks writing", () => {
  it("writes column breaks with correct XML structure", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["A"]],
      colBreaks: [4], // 0-based: break after column E
    };

    const xml = writeXml(sheet);
    const doc = parseXml(xml);

    const cb = findChild(doc, "colBreaks");
    expect(cb).toBeDefined();
    expect(cb.attrs["count"]).toBe("1");
    expect(cb.attrs["manualBreakCount"]).toBe("1");

    const brks = findChildren(cb, "brk");
    expect(brks).toHaveLength(1);

    expect(brks[0].attrs["id"]).toBe("5"); // 0-based + 1
    expect(brks[0].attrs["max"]).toBe("1048575");
    expect(brks[0].attrs["man"]).toBe("1");
  });
});

// ── Both Row and Column Breaks ──────────────────────────────────────

describe("page breaks — both row and column breaks", () => {
  it("writes both row and column breaks", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["A"]],
      rowBreaks: [9],
      colBreaks: [4],
    };

    const xml = writeXml(sheet);
    const doc = parseXml(xml);

    expect(findChild(doc, "rowBreaks")).toBeDefined();
    expect(findChild(doc, "colBreaks")).toBeDefined();
  });
});

// ── No Breaks ───────────────────────────────────────────────────────

describe("page breaks — no breaks", () => {
  it("omits rowBreaks/colBreaks elements when no breaks defined", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["A"]],
    };

    const xml = writeXml(sheet);
    const doc = parseXml(xml);

    expect(findChild(doc, "rowBreaks")).toBeUndefined();
    expect(findChild(doc, "colBreaks")).toBeUndefined();
  });

  it("omits elements for empty arrays", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["A"]],
      rowBreaks: [],
      colBreaks: [],
    };

    const xml = writeXml(sheet);
    const doc = parseXml(xml);

    expect(findChild(doc, "rowBreaks")).toBeUndefined();
    expect(findChild(doc, "colBreaks")).toBeUndefined();
  });
});

// ── Round-trip (write → read) ───────────────────────────────────────

describe("page breaks — round-trip", () => {
  it("row breaks survive write → read cycle", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Row 1"], ["Row 2"], ["Row 3"]],
          rowBreaks: [0, 1],
        },
      ],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].rowBreaks).toBeDefined();
    expect(wb.sheets[0].rowBreaks).toEqual([0, 1]);
  });

  it("column breaks survive write → read cycle", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["A", "B", "C"]],
          colBreaks: [1],
        },
      ],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].colBreaks).toBeDefined();
    expect(wb.sheets[0].colBreaks).toEqual([1]);
  });

  it("both row and column breaks survive round-trip", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["A"]],
          rowBreaks: [9, 24],
          colBreaks: [4, 7],
        },
      ],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].rowBreaks).toEqual([9, 24]);
    expect(wb.sheets[0].colBreaks).toEqual([4, 7]);
  });

  it("multiple breaks are sorted after reading", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["A"]],
          rowBreaks: [24, 9, 15],
        },
      ],
    });

    const wb = await readXlsx(xlsx);
    // Should come back sorted
    expect(wb.sheets[0].rowBreaks).toEqual([9, 15, 24]);
  });
});
