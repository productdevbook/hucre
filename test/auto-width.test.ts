import { describe, it, expect } from "vitest";
import { calculateColumnWidth, measureValueWidth } from "../src/xlsx/auto-width";
import { writeXlsx } from "../src/xlsx/writer";
import { ZipReader } from "../src/zip/reader";
import { parseXml } from "../src/xml/parser";
import type { CellValue } from "../src/_types";

// ── Helpers ──────────────────────────────────────────────────────────

const decoder = new TextDecoder("utf-8");

async function extractXml(data: Uint8Array, path: string): Promise<string> {
  const zip = new ZipReader(data);
  const raw = await zip.extract(path);
  return decoder.decode(raw);
}

async function parseXmlFromZip(data: Uint8Array, path: string) {
  const xml = await extractXml(data, path);
  return parseXml(xml);
}

function findChild(el: { children: Array<unknown> }, localName: string): any {
  return el.children.find((c: any) => typeof c !== "string" && (c.local || c.tag) === localName);
}

function findChildren(el: { children: Array<unknown> }, localName: string): any[] {
  return el.children.filter((c: any) => typeof c !== "string" && (c.local || c.tag) === localName);
}

// ── measureValueWidth ────────────────────────────────────────────────

describe("measureValueWidth", () => {
  it("returns 0 for null", () => {
    expect(measureValueWidth(null)).toBe(0);
  });

  it("returns 0 for empty string", () => {
    expect(measureValueWidth("")).toBe(0);
  });

  it("measures short string", () => {
    expect(measureValueWidth("Hi")).toBe(2);
  });

  it("measures longer string", () => {
    expect(measureValueWidth("A very long product name")).toBe(24);
  });

  it("returns 5 for boolean true (TRUE)", () => {
    expect(measureValueWidth(true)).toBe(5);
  });

  it("returns 6 for boolean false (FALSE)", () => {
    expect(measureValueWidth(false)).toBe(6);
  });

  it("measures number without format", () => {
    // "9.99" = 4 chars
    expect(measureValueWidth(9.99)).toBe(4);
  });

  it("measures larger number without format", () => {
    // "1234567.89" = 10 chars
    expect(measureValueWidth(1234567.89)).toBe(10);
  });

  it("measures number with currency format", () => {
    // "$#,##0.00" applied to 1234567.89 should produce "1,234,567.89" (12 chars) + "$" (1 literal)
    const width = measureValueWidth(1234567.89, "$#,##0.00");
    expect(width).toBe(13);
  });

  it("measures number with simple decimal format", () => {
    // "#,##0.00" applied to 9.99 -> "9.99" (4 chars)
    const width = measureValueWidth(9.99, "#,##0.00");
    expect(width).toBe(4);
  });

  it("measures date without format", () => {
    const date = new Date(Date.UTC(2024, 0, 15)); // 2024-01-15
    // "2024-01-15" = 10 chars
    expect(measureValueWidth(date)).toBe(10);
  });

  it("measures date with format", () => {
    const date = new Date(Date.UTC(2024, 0, 15));
    // "yyyy-mm-dd" -> "2024-01-15" = 10 chars
    expect(measureValueWidth(date, "yyyy-mm-dd")).toBe(10);
  });

  it("measures CJK characters as double width", () => {
    // 3 CJK characters = 6 width units
    expect(measureValueWidth("\u4F60\u597D\u554A")).toBe(6);
  });

  it("measures mixed ASCII and CJK", () => {
    // "Hi" (2) + "\u4F60" (2) = 4
    expect(measureValueWidth("Hi\u4F60")).toBe(4);
  });

  it("handles multiline strings - takes longest line", () => {
    const value = "Short\nA much longer line\nMed";
    // "A much longer line" = 18 chars (the longest)
    expect(measureValueWidth(value)).toBe(18);
  });

  it("handles multiline with CJK", () => {
    // line1: "AB" = 2, line2: "\u4F60\u597D" = 4
    const value = "AB\n\u4F60\u597D";
    expect(measureValueWidth(value)).toBe(4);
  });
});

// ── calculateColumnWidth ─────────────────────────────────────────────

describe("calculateColumnWidth", () => {
  it("returns min width for empty values", () => {
    const width = calculateColumnWidth([null, null]);
    expect(width).toBe(8); // default min
  });

  it("returns min width for short content", () => {
    // "Hi" = 2 chars * 1.1 multiplier + 2 padding = 4.2 -> rounds to 4.5
    // But min is 8, so should be 8
    const width = calculateColumnWidth(["Hi"]);
    expect(width).toBe(8);
  });

  it("calculates width for longer content", () => {
    const width = calculateColumnWidth(["A very long product name"]);
    // 24 * 1.1 + 2 = 28.4 -> rounds to 28.5
    expect(width).toBe(28.5);
  });

  it("takes the max width across all values", () => {
    const values: CellValue[] = ["Short", "A much longer value here", "Med"];
    const width = calculateColumnWidth(values);
    // "A much longer value here" = 24 chars * 1.1 + 2 = 28.4 -> 28.5
    expect(width).toBe(28.5);
  });

  it("respects minWidth option", () => {
    const width = calculateColumnWidth(["Hi"], { minWidth: 15 });
    expect(width).toBe(15);
  });

  it("respects maxWidth option", () => {
    const longStr = "A".repeat(300);
    const width = calculateColumnWidth([longStr], { maxWidth: 50 });
    expect(width).toBe(50);
  });

  it("respects the Excel maximum of 255", () => {
    const longStr = "A".repeat(500);
    const width = calculateColumnWidth([longStr]);
    expect(width).toBe(255);
  });

  it("applies bold multiplier", () => {
    const normalWidth = calculateColumnWidth(["Some text here"]);
    const boldWidth = calculateColumnWidth(["Some text here"], {
      font: { bold: true },
    });
    expect(boldWidth).toBeGreaterThan(normalWidth);
  });

  it("uses numFmt for number formatting", () => {
    // With format, 1234567.89 -> "1,234,567.89" + "$" = 13 chars
    // Without format, 1234567.89 -> "1234567.89" = 10 chars
    const withFmt = calculateColumnWidth([1234567.89], {
      numFmt: "$#,##0.00",
    });
    const withoutFmt = calculateColumnWidth([1234567.89]);
    expect(withFmt).toBeGreaterThan(withoutFmt);
  });

  it("handles mixed column (header + data)", () => {
    const values: CellValue[] = [
      "Product Name",
      "Widget A",
      "Super Deluxe Widget B Extended",
      "Gadget",
    ];
    const width = calculateColumnWidth(values);
    // "Super Deluxe Widget B Extended" = 30 chars -> 30 * 1.1 + 2 = 35 -> 35
    expect(width).toBe(35);
  });

  it("handles custom padding", () => {
    const defaultPad = calculateColumnWidth(["Some text here"]);
    const noPad = calculateColumnWidth(["Some text here"], { padding: 0 });
    const bigPad = calculateColumnWidth(["Some text here"], { padding: 5 });
    expect(noPad).toBeLessThan(defaultPad);
    expect(bigPad).toBeGreaterThan(defaultPad);
  });

  it("handles percentage format", () => {
    // 0.1234 with "0.00%" -> "12.34%" = 6 chars
    const width = measureValueWidth(0.1234, "0.00%");
    expect(width).toBe(6);
  });
});

// ── Integration: writeXlsx with autoWidth ────────────────────────────

describe("writeXlsx with autoWidth", () => {
  it("generates col element with calculated width for autoWidth column", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          columns: [
            { header: "ID", key: "id", autoWidth: true },
            {
              header: "Product Name",
              key: "name",
              autoWidth: true,
            },
            { header: "Price", key: "price", width: 12 },
          ],
          data: [
            { id: 1, name: "Widget Alpha", price: 9.99 },
            { id: 2, name: "Super Long Product Name Here", price: 19.99 },
            { id: 3, name: "Gizmo", price: 4.5 },
          ],
        },
      ],
    });

    const root = await parseXmlFromZip(data, "xl/worksheets/sheet1.xml");
    const cols = findChild(root, "cols");
    expect(cols).toBeDefined();

    const colDefs = findChildren(cols, "col");

    // Column A (ID) should have autoWidth calculated
    const colA = colDefs.find((c: any) => c.attrs?.min === "1");
    expect(colA).toBeDefined();
    expect(colA.attrs?.customWidth).toBe("true");
    const widthA = parseFloat(colA.attrs?.width);
    expect(widthA).toBeGreaterThan(0);

    // Column B (Product Name) should have autoWidth calculated
    const colB = colDefs.find((c: any) => c.attrs?.min === "2");
    expect(colB).toBeDefined();
    expect(colB.attrs?.customWidth).toBe("true");
    const widthB = parseFloat(colB.attrs?.width);
    // "Super Long Product Name Here" is 28 chars, wider than "Product Name" (12 chars)
    // So the width should be based on the longest value
    expect(widthB).toBeGreaterThan(widthA);

    // Column C (Price) should have explicit width=12
    const colC = colDefs.find((c: any) => c.attrs?.min === "3");
    expect(colC).toBeDefined();
    expect(parseFloat(colC.attrs?.width)).toBe(12);
  });

  it("explicit width takes precedence over autoWidth", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          columns: [
            {
              header: "Name",
              key: "name",
              width: 20,
              autoWidth: true,
            },
          ],
          data: [{ name: "A very long name that would need more than 20 chars" }],
        },
      ],
    });

    const root = await parseXmlFromZip(data, "xl/worksheets/sheet1.xml");
    const cols = findChild(root, "cols");
    const colDefs = findChildren(cols, "col");
    const colA = colDefs[0];
    // Explicit width should win
    expect(parseFloat(colA.attrs?.width)).toBe(20);
  });

  it("autoWidth with rows-based data", async () => {
    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          columns: [{ autoWidth: true }, { autoWidth: true }],
          rows: [
            ["Header A", "Header B"],
            ["Short", "A much longer cell value here"],
          ],
        },
      ],
    });

    const root = await parseXmlFromZip(data, "xl/worksheets/sheet1.xml");
    const cols = findChild(root, "cols");
    expect(cols).toBeDefined();

    const colDefs = findChildren(cols, "col");
    expect(colDefs.length).toBe(2);

    const widthA = parseFloat(colDefs[0].attrs?.width);
    const widthB = parseFloat(colDefs[1].attrs?.width);
    // Column B has longer content, should be wider
    expect(widthB).toBeGreaterThan(widthA);
  });
});
