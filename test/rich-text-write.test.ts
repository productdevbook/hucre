import { describe, it, expect } from "vitest";
import { parseXml } from "../src/xml/parser";
import { writeXlsx } from "../src/xlsx/writer";
import { readXlsx } from "../src/xlsx/reader";
import { createStylesCollector } from "../src/xlsx/styles-writer";
import { createSharedStrings, writeWorksheetXml } from "../src/xlsx/worksheet-writer";
import type { WriteSheet, Cell } from "../src/_types";

// ── Helpers ──────────────────────────────────────────────────────────

function findChild(el: { children: Array<unknown> }, localName: string): any {
  return el.children.find((c: any) => typeof c !== "string" && (c.local || c.tag) === localName);
}

function findChildren(el: { children: Array<unknown> }, localName: string): any[] {
  return el.children.filter((c: any) => typeof c !== "string" && (c.local || c.tag) === localName);
}

function getElementText(el: { children: Array<unknown> }): string {
  return el.children.filter((c: unknown) => typeof c === "string").join("");
}

function writeXml(sheet: WriteSheet): string {
  const styles = createStylesCollector();
  const ss = createSharedStrings();
  const result = writeWorksheetXml(sheet, styles, ss);
  return result.xml;
}

function parseSheet(xml: string) {
  return parseXml(xml);
}

// ── Rich Text Inline String Writing Tests ────────────────────────────

describe("rich text write — basic", () => {
  it("writes cell with rich text as inlineStr", () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      richText: [{ text: "Bold", font: { bold: true } }, { text: " Normal" }],
    });

    const sheet: WriteSheet = {
      name: "Test",
      rows: [[""]],
      cells,
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetData = findChild(doc, "sheetData");
    const row = findChild(sheetData, "row");
    const c = findChild(row, "c");

    expect(c.attrs["t"]).toBe("inlineStr");
    expect(c.attrs["r"]).toBe("A1");

    const is = findChild(c, "is");
    expect(is).toBeDefined();

    const runs = findChildren(is, "r");
    expect(runs.length).toBe(2);

    // First run: Bold
    const rPr1 = findChild(runs[0], "rPr");
    expect(rPr1).toBeDefined();
    const b = findChild(rPr1, "b");
    expect(b).toBeDefined();
    const t1 = findChild(runs[0], "t");
    expect(getElementText(t1)).toBe("Bold");

    // Second run: Normal (no rPr)
    const rPr2 = findChild(runs[1], "rPr");
    expect(rPr2).toBeUndefined();
    const t2 = findChild(runs[1], "t");
    expect(getElementText(t2)).toBe(" Normal");
  });
});

describe("rich text write — font properties", () => {
  it("writes rich text with color", () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      richText: [{ text: "Red", font: { color: { rgb: "FF0000" } } }],
    });

    const sheet: WriteSheet = {
      name: "Test",
      rows: [[""]],
      cells,
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetData = findChild(doc, "sheetData");
    const row = findChild(sheetData, "row");
    const c = findChild(row, "c");
    const is = findChild(c, "is");
    const r = findChild(is, "r");
    const rPr = findChild(r, "rPr");
    const color = findChild(rPr, "color");
    expect(color).toBeDefined();
    expect(color.attrs["rgb"]).toBe("FFFF0000");
  });

  it("writes rich text with italic", () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      richText: [{ text: "Italic", font: { italic: true } }],
    });

    const sheet: WriteSheet = {
      name: "Test",
      rows: [[""]],
      cells,
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetData = findChild(doc, "sheetData");
    const row = findChild(sheetData, "row");
    const c = findChild(row, "c");
    const is = findChild(c, "is");
    const r = findChild(is, "r");
    const rPr = findChild(r, "rPr");
    const i = findChild(rPr, "i");
    expect(i).toBeDefined();
  });

  it("writes rich text with underline", () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      richText: [{ text: "Underline", font: { underline: true } }],
    });

    const sheet: WriteSheet = {
      name: "Test",
      rows: [[""]],
      cells,
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetData = findChild(doc, "sheetData");
    const row = findChild(sheetData, "row");
    const c = findChild(row, "c");
    const is = findChild(c, "is");
    const r = findChild(is, "r");
    const rPr = findChild(r, "rPr");
    const u = findChild(rPr, "u");
    expect(u).toBeDefined();
  });

  it("writes rich text with double underline", () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      richText: [{ text: "Double", font: { underline: "double" } }],
    });

    const sheet: WriteSheet = {
      name: "Test",
      rows: [[""]],
      cells,
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetData = findChild(doc, "sheetData");
    const row = findChild(sheetData, "row");
    const c = findChild(row, "c");
    const is = findChild(c, "is");
    const r = findChild(is, "r");
    const rPr = findChild(r, "rPr");
    const u = findChild(rPr, "u");
    expect(u).toBeDefined();
    expect(u.attrs["val"]).toBe("double");
  });

  it("writes rich text with strikethrough", () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      richText: [{ text: "Strike", font: { strikethrough: true } }],
    });

    const sheet: WriteSheet = {
      name: "Test",
      rows: [[""]],
      cells,
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetData = findChild(doc, "sheetData");
    const row = findChild(sheetData, "row");
    const c = findChild(row, "c");
    const is = findChild(c, "is");
    const r = findChild(is, "r");
    const rPr = findChild(r, "rPr");
    const strike = findChild(rPr, "strike");
    expect(strike).toBeDefined();
  });

  it("writes rich text with font name and size", () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      richText: [{ text: "Styled", font: { name: "Arial", size: 14 } }],
    });

    const sheet: WriteSheet = {
      name: "Test",
      rows: [[""]],
      cells,
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetData = findChild(doc, "sheetData");
    const row = findChild(sheetData, "row");
    const c = findChild(row, "c");
    const is = findChild(c, "is");
    const r = findChild(is, "r");
    const rPr = findChild(r, "rPr");
    const sz = findChild(rPr, "sz");
    expect(sz.attrs["val"]).toBe("14");
    const rFont = findChild(rPr, "rFont");
    expect(rFont.attrs["val"]).toBe("Arial");
  });

  it("writes rich text with multiple font properties", () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      richText: [
        {
          text: "Fancy",
          font: {
            bold: true,
            italic: true,
            size: 16,
            color: { rgb: "0000FF" },
            name: "Calibri",
          },
        },
      ],
    });

    const sheet: WriteSheet = {
      name: "Test",
      rows: [[""]],
      cells,
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetData = findChild(doc, "sheetData");
    const row = findChild(sheetData, "row");
    const c = findChild(row, "c");
    const is = findChild(c, "is");
    const r = findChild(is, "r");
    const rPr = findChild(r, "rPr");
    expect(findChild(rPr, "b")).toBeDefined();
    expect(findChild(rPr, "i")).toBeDefined();
    expect(findChild(rPr, "sz").attrs["val"]).toBe("16");
    expect(findChild(rPr, "color").attrs["rgb"]).toBe("FF0000FF");
    expect(findChild(rPr, "rFont").attrs["val"]).toBe("Calibri");
  });
});

// ── Rich Text with Empty Run Text ────────────────────────────────────

describe("rich text write — empty run text", () => {
  it("writes rich text with empty run text", () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      richText: [{ text: "", font: { bold: true } }, { text: "After empty" }],
    });

    const sheet: WriteSheet = {
      name: "Test",
      rows: [[""]],
      cells,
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetData = findChild(doc, "sheetData");
    const row = findChild(sheetData, "row");
    const c = findChild(row, "c");
    expect(c.attrs["t"]).toBe("inlineStr");

    const is = findChild(c, "is");
    const runs = findChildren(is, "r");
    expect(runs.length).toBe(2);
  });
});

// ── Mixed Sheet: Plain String + Rich Text ─────────────────────────────

describe("rich text write — mixed sheet", () => {
  it("writes some cells as plain string and some as rich text", () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("1,0", {
      richText: [{ text: "Bold", font: { bold: true } }, { text: " text" }],
    });

    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Plain string"], ["placeholder"], [42]],
      cells,
    };

    const xml = writeXml(sheet);
    const doc = parseSheet(xml);
    const sheetData = findChild(doc, "sheetData");
    const rows = findChildren(sheetData, "row");

    // Row 1: plain string (shared string type 's')
    const c1 = findChild(rows[0], "c");
    expect(c1.attrs["t"]).toBe("s");

    // Row 2: rich text (inlineStr)
    const c2 = findChild(rows[1], "c");
    expect(c2.attrs["t"]).toBe("inlineStr");

    // Row 3: number
    const c3 = findChild(rows[2], "c");
    expect(c3.attrs["t"]).toBeUndefined(); // numbers don't have type attr
  });
});

// ── Round-trip Tests ────────────────────────────────────────────────

describe("rich text write — round-trip", () => {
  it("round-trips bold + normal rich text", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      richText: [
        { text: "Bold", font: { bold: true, name: "Calibri", size: 11 } },
        { text: " Normal" },
      ],
    });

    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [[""]],
          cells,
        },
      ],
    });

    const workbook = await readXlsx(data);
    const cell = workbook.sheets[0].cells?.get("0,0");
    expect(cell).toBeDefined();
    expect(cell!.type).toBe("richText");
    expect(cell!.richText).toBeDefined();
    expect(cell!.richText!.length).toBe(2);
    expect(cell!.richText![0].text).toBe("Bold");
    expect(cell!.richText![0].font?.bold).toBe(true);
    expect(cell!.richText![1].text).toBe(" Normal");
  });

  it("round-trips rich text with color", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      richText: [
        { text: "Red", font: { color: { rgb: "FF0000" }, size: 11 } },
        { text: " Blue", font: { color: { rgb: "0000FF" }, size: 11 } },
      ],
    });

    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [[""]],
          cells,
        },
      ],
    });

    const workbook = await readXlsx(data);
    const cell = workbook.sheets[0].cells?.get("0,0");
    expect(cell).toBeDefined();
    expect(cell!.richText).toBeDefined();
    expect(cell!.richText!.length).toBe(2);
    expect(cell!.richText![0].text).toBe("Red");
    expect(cell!.richText![0].font?.color?.rgb).toBe("FF0000");
    expect(cell!.richText![1].text).toBe(" Blue");
    expect(cell!.richText![1].font?.color?.rgb).toBe("0000FF");
  });

  it("round-trips rich text with italic and underline", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      richText: [
        { text: "Italic", font: { italic: true } },
        { text: " Underline", font: { underline: true } },
      ],
    });

    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [[""]],
          cells,
        },
      ],
    });

    const workbook = await readXlsx(data);
    const cell = workbook.sheets[0].cells?.get("0,0");
    expect(cell).toBeDefined();
    expect(cell!.richText).toBeDefined();
    expect(cell!.richText![0].font?.italic).toBe(true);
    expect(cell!.richText![1].font?.underline).toBe(true);
  });

  it("round-trips mixed sheet with plain and rich text", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("1,0", {
      richText: [{ text: "Bold", font: { bold: true } }, { text: " text" }],
    });

    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Plain"], [""]],
          cells,
        },
      ],
    });

    const workbook = await readXlsx(data);

    // Row 0: plain string
    expect(workbook.sheets[0].rows[0][0]).toBe("Plain");

    // Row 1: rich text
    const cell = workbook.sheets[0].cells?.get("1,0");
    expect(cell).toBeDefined();
    expect(cell!.type).toBe("richText");
    expect(cell!.richText!.length).toBe(2);
    expect(cell!.richText![0].text).toBe("Bold");
    expect(cell!.richText![0].font?.bold).toBe(true);
    expect(cell!.richText![1].text).toBe(" text");
  });

  it("round-trips value from rich text runs", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      richText: [{ text: "Hello " }, { text: "World" }],
    });

    const data = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [[""]],
          cells,
        },
      ],
    });

    const workbook = await readXlsx(data);
    // The plain value should be the concatenation of all run texts
    expect(workbook.sheets[0].rows[0][0]).toBe("Hello World");
  });
});
