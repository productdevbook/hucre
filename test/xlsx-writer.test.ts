import { describe, it, expect } from "vitest";
import { ZipReader } from "../src/zip/reader";
import { parseXml } from "../src/xml/parser";
import { writeXlsx } from "../src/xlsx/writer";
import { writeContentTypes } from "../src/xlsx/content-types-writer";
import { writeWorkbookXml, writeWorkbookRels, writeRootRels } from "../src/xlsx/workbook-writer";
import { createStylesCollector } from "../src/xlsx/styles-writer";
import {
  createSharedStrings,
  writeSharedStringsXml,
  writeWorksheetXml,
  colToLetter,
  cellRef,
} from "../src/xlsx/worksheet-writer";
import type { WriteSheet, CellStyle, CellValue } from "../src/_types";

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

function getElementText(el: { children: Array<unknown> }): string {
  return el.children.filter((c: unknown) => typeof c === "string").join("");
}

// ── colToLetter / cellRef ────────────────────────────────────────────

describe("colToLetter", () => {
  it("converts 0 to A", () => {
    expect(colToLetter(0)).toBe("A");
  });

  it("converts 25 to Z", () => {
    expect(colToLetter(25)).toBe("Z");
  });

  it("converts 26 to AA", () => {
    expect(colToLetter(26)).toBe("AA");
  });

  it("converts 701 to ZZ", () => {
    expect(colToLetter(701)).toBe("ZZ");
  });

  it("converts 702 to AAA", () => {
    expect(colToLetter(702)).toBe("AAA");
  });
});

describe("cellRef", () => {
  it("returns A1 for (0, 0)", () => {
    expect(cellRef(0, 0)).toBe("A1");
  });

  it("returns C3 for (2, 2)", () => {
    expect(cellRef(2, 2)).toBe("C3");
  });

  it("returns AA10 for (9, 26)", () => {
    expect(cellRef(9, 26)).toBe("AA10");
  });
});

// ── Content Types Writer ─────────────────────────────────────────────

describe("writeContentTypes", () => {
  it("generates valid content types XML", () => {
    const xml = writeContentTypes(2, true);
    const doc = parseXml(xml);

    // Check it has Types root element
    expect(doc.local || doc.tag).toBe("Types");

    // Check defaults
    const defaults = findChildren(doc, "Default");
    expect(defaults.length).toBe(2);

    const relsDefault = defaults.find((d: any) => d.attrs["Extension"] === "rels");
    expect(relsDefault).toBeDefined();
    expect(relsDefault.attrs["ContentType"]).toContain("relationships");

    const xmlDefault = defaults.find((d: any) => d.attrs["Extension"] === "xml");
    expect(xmlDefault).toBeDefined();
    expect(xmlDefault.attrs["ContentType"]).toBe("application/xml");

    // Check overrides
    const overrides = findChildren(doc, "Override");
    // workbook + 2 sheets + styles + theme + sharedStrings = 6
    expect(overrides.length).toBe(6);

    const workbookOverride = overrides.find((o: any) => o.attrs["PartName"] === "/xl/workbook.xml");
    expect(workbookOverride).toBeDefined();

    const sheet1Override = overrides.find(
      (o: any) => o.attrs["PartName"] === "/xl/worksheets/sheet1.xml",
    );
    expect(sheet1Override).toBeDefined();

    const sheet2Override = overrides.find(
      (o: any) => o.attrs["PartName"] === "/xl/worksheets/sheet2.xml",
    );
    expect(sheet2Override).toBeDefined();
  });

  it("omits shared strings override when not needed", () => {
    const xml = writeContentTypes(1, false);
    const doc = parseXml(xml);

    const overrides = findChildren(doc, "Override");
    const ssOverride = overrides.find((o: any) => o.attrs["PartName"] === "/xl/sharedStrings.xml");
    expect(ssOverride).toBeUndefined();
  });
});

// ── Workbook Writer ──────────────────────────────────────────────────

describe("writeWorkbookXml", () => {
  it("generates workbook with correct sheet names", () => {
    const sheets: WriteSheet[] = [{ name: "Sheet1" }, { name: "Data" }];
    const xml = writeWorkbookXml(sheets);
    const doc = parseXml(xml);

    const sheetsEl = findChild(doc, "sheets");
    expect(sheetsEl).toBeDefined();

    const sheetEls = findChildren(sheetsEl, "sheet");
    expect(sheetEls.length).toBe(2);
    expect(sheetEls[0].attrs["name"]).toBe("Sheet1");
    expect(sheetEls[0].attrs["sheetId"]).toBe("1");
    expect(sheetEls[0].attrs["r:id"]).toBe("rId1");
    expect(sheetEls[1].attrs["name"]).toBe("Data");
    expect(sheetEls[1].attrs["sheetId"]).toBe("2");
  });

  it("marks hidden sheets", () => {
    const sheets: WriteSheet[] = [
      { name: "Visible" },
      { name: "Hidden", hidden: true },
      { name: "VeryHidden", veryHidden: true },
    ];
    const xml = writeWorkbookXml(sheets);
    const doc = parseXml(xml);

    const sheetsEl = findChild(doc, "sheets");
    const sheetEls = findChildren(sheetsEl, "sheet");
    expect(sheetEls[0].attrs["state"]).toBeUndefined();
    expect(sheetEls[1].attrs["state"]).toBe("hidden");
    expect(sheetEls[2].attrs["state"]).toBe("veryHidden");
  });
});

describe("writeWorkbookRels", () => {
  it("generates correct relationships", () => {
    const xml = writeWorkbookRels(2, true);
    const doc = parseXml(xml);

    const rels = findChildren(doc, "Relationship");
    // 2 worksheets + 1 styles + 1 shared strings + 1 theme = 5
    expect(rels.length).toBe(5);

    // First two should be worksheets
    expect(rels[0].attrs["Target"]).toBe("worksheets/sheet1.xml");
    expect(rels[1].attrs["Target"]).toBe("worksheets/sheet2.xml");

    // Styles
    expect(rels[2].attrs["Target"]).toBe("styles.xml");

    // Shared strings
    expect(rels[3].attrs["Target"]).toBe("sharedStrings.xml");

    // Theme
    expect(rels[4].attrs["Target"]).toBe("theme/theme1.xml");
  });

  it("omits shared strings rel when not present", () => {
    const xml = writeWorkbookRels(1, false);
    const doc = parseXml(xml);

    const rels = findChildren(doc, "Relationship");
    // 1 worksheet + 1 styles + 1 theme = 3
    expect(rels.length).toBe(3);
  });
});

describe("writeRootRels", () => {
  it("generates root rels pointing to workbook", () => {
    const xml = writeRootRels();
    const doc = parseXml(xml);

    const rels = findChildren(doc, "Relationship");
    expect(rels.length).toBe(1);
    expect(rels[0].attrs["Target"]).toBe("xl/workbook.xml");
    expect(rels[0].attrs["Type"]).toContain("officeDocument");
  });
});

// ── Shared Strings ───────────────────────────────────────────────────

describe("SharedStringsCollector", () => {
  it("deduplicates strings", () => {
    const ss = createSharedStrings();
    const idx1 = ss.add("hello");
    const idx2 = ss.add("world");
    const idx3 = ss.add("hello");

    expect(idx1).toBe(0);
    expect(idx2).toBe(1);
    expect(idx3).toBe(0); // deduplicated
    expect(ss.count()).toBe(2);
    expect(ss.getAll()).toEqual(["hello", "world"]);
  });

  it("generates shared strings XML", () => {
    const ss = createSharedStrings();
    ss.add("Hello");
    ss.add("World");

    const xml = writeSharedStringsXml(ss);
    const doc = parseXml(xml);

    expect(doc.local || doc.tag).toBe("sst");
    expect(doc.attrs["count"]).toBe("2");
    expect(doc.attrs["uniqueCount"]).toBe("2");

    const siElements = findChildren(doc, "si");
    expect(siElements.length).toBe(2);

    const t1 = findChild(siElements[0], "t");
    expect(t1.text || getElementText(t1)).toBe("Hello");

    const t2 = findChild(siElements[1], "t");
    expect(t2.text || getElementText(t2)).toBe("World");
  });
});

// ── Styles Writer ────────────────────────────────────────────────────

describe("StylesCollector", () => {
  it("generates styles.xml with default entries", () => {
    const styles = createStylesCollector();
    const xml = styles.toXml();
    const doc = parseXml(xml);

    expect(doc.local || doc.tag).toBe("styleSheet");

    // Should have default font
    const fontsEl = findChild(doc, "fonts");
    expect(fontsEl).toBeDefined();
    expect(fontsEl.attrs["count"]).toBe("1");

    // Should have 2 default fills (none + gray125)
    const fillsEl = findChild(doc, "fills");
    expect(fillsEl).toBeDefined();
    expect(fillsEl.attrs["count"]).toBe("2");

    // Should have 1 default border
    const bordersEl = findChild(doc, "borders");
    expect(bordersEl).toBeDefined();
    expect(bordersEl.attrs["count"]).toBe("1");

    // Should have 1 cellXf (default)
    const cellXfsEl = findChild(doc, "cellXfs");
    expect(cellXfsEl).toBeDefined();
    expect(cellXfsEl.attrs["count"]).toBe("1");
  });

  it("deduplicates identical styles", () => {
    const styles = createStylesCollector();
    const style: CellStyle = { font: { bold: true } };
    const idx1 = styles.addStyle(style);
    const idx2 = styles.addStyle(style);

    expect(idx1).toBe(idx2);
    expect(idx1).toBeGreaterThan(0); // not the default

    const xml = styles.toXml();
    const doc = parseXml(xml);

    const cellXfsEl = findChild(doc, "cellXfs");
    // 1 default + 1 bold = 2
    expect(cellXfsEl.attrs["count"]).toBe("2");
  });

  it("handles custom number formats", () => {
    const styles = createStylesCollector();
    const id = styles.addNumFmt("yyyy-mm-dd");

    expect(id).toBe(164); // first custom ID

    const xml = styles.toXml();
    const doc = parseXml(xml);

    const numFmtsEl = findChild(doc, "numFmts");
    expect(numFmtsEl).toBeDefined();
    expect(numFmtsEl.attrs["count"]).toBe("1");

    const numFmt = findChild(numFmtsEl, "numFmt");
    expect(numFmt.attrs["numFmtId"]).toBe("164");
    expect(numFmt.attrs["formatCode"]).toBe("yyyy-mm-dd");
  });

  it("adds fonts with all properties", () => {
    const styles = createStylesCollector();
    styles.addStyle({
      font: {
        name: "Arial",
        size: 14,
        bold: true,
        italic: true,
        underline: true,
        strikethrough: true,
        color: { rgb: "FF0000" },
      },
    });

    const xml = styles.toXml();
    const doc = parseXml(xml);

    const fontsEl = findChild(doc, "fonts");
    const fontEls = findChildren(fontsEl, "font");
    // default + custom = 2
    expect(fontEls.length).toBe(2);

    const customFont = fontEls[1];
    expect(findChild(customFont, "b")).toBeDefined();
    expect(findChild(customFont, "i")).toBeDefined();
    expect(findChild(customFont, "u")).toBeDefined();
    expect(findChild(customFont, "strike")).toBeDefined();
    expect(findChild(customFont, "sz").attrs["val"]).toBe("14");
    expect(findChild(customFont, "name").attrs["val"]).toBe("Arial");

    const colorEl = findChild(customFont, "color");
    expect(colorEl.attrs["rgb"]).toBe("FFFF0000"); // prefixed with FF
  });

  it("adds fills correctly", () => {
    const styles = createStylesCollector();
    styles.addStyle({
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { rgb: "FFFF00" },
      },
    });

    const xml = styles.toXml();
    const doc = parseXml(xml);

    const fillsEl = findChild(doc, "fills");
    const fillEls = findChildren(fillsEl, "fill");
    // 2 defaults + 1 custom = 3
    expect(fillEls.length).toBe(3);
  });

  it("adds borders correctly", () => {
    const styles = createStylesCollector();
    styles.addStyle({
      border: {
        top: { style: "thin", color: { rgb: "000000" } },
        bottom: { style: "thin", color: { rgb: "000000" } },
        left: { style: "thin", color: { rgb: "000000" } },
        right: { style: "thin", color: { rgb: "000000" } },
      },
    });

    const xml = styles.toXml();
    const doc = parseXml(xml);

    const bordersEl = findChild(doc, "borders");
    const borderEls = findChildren(bordersEl, "border");
    // 1 default + 1 custom = 2
    expect(borderEls.length).toBe(2);

    const customBorder = borderEls[1];
    const leftEl = findChild(customBorder, "left");
    expect(leftEl.attrs["style"]).toBe("thin");
  });

  it("uses default font when specified", () => {
    const styles = createStylesCollector({ name: "Times New Roman", size: 12 });
    const xml = styles.toXml();
    const doc = parseXml(xml);

    const fontsEl = findChild(doc, "fonts");
    const fontEls = findChildren(fontsEl, "font");
    const nameEl = findChild(fontEls[0], "name");
    expect(nameEl.attrs["val"]).toBe("Times New Roman");
    const sizeEl = findChild(fontEls[0], "sz");
    expect(sizeEl.attrs["val"]).toBe("12");
  });
});

// ── Worksheet Writer ─────────────────────────────────────────────────

describe("writeWorksheetXml", () => {
  it("generates worksheet with string/number/boolean cells", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [
        ["Hello", 42, true],
        ["World", 3.14, false],
      ],
    };
    const styles = createStylesCollector();
    const ss = createSharedStrings();

    const xml = writeWorksheetXml(sheet, styles, ss).xml;
    const doc = parseXml(xml);

    const sheetData = findChild(doc, "sheetData");
    expect(sheetData).toBeDefined();

    const rows = findChildren(sheetData, "row");
    expect(rows.length).toBe(2);

    // Row 1
    expect(rows[0].attrs["r"]).toBe("1");
    const row1Cells = findChildren(rows[0], "c");
    expect(row1Cells.length).toBe(3);

    // A1 = "Hello" (shared string)
    expect(row1Cells[0].attrs["r"]).toBe("A1");
    expect(row1Cells[0].attrs["t"]).toBe("s");
    const a1v = findChild(row1Cells[0], "v");
    expect(a1v.text || getElementText(a1v)).toBe("0"); // index 0

    // B1 = 42 (number)
    expect(row1Cells[1].attrs["r"]).toBe("B1");
    expect(row1Cells[1].attrs["t"]).toBeUndefined(); // numbers have no t attribute
    const b1v = findChild(row1Cells[1], "v");
    expect(b1v.text || getElementText(b1v)).toBe("42");

    // C1 = true (boolean)
    expect(row1Cells[2].attrs["r"]).toBe("C1");
    expect(row1Cells[2].attrs["t"]).toBe("b");
    const c1v = findChild(row1Cells[2], "v");
    expect(c1v.text || getElementText(c1v)).toBe("1");

    // Row 2
    const row2Cells = findChildren(rows[1], "c");
    // D2 = false
    const c2v = findChild(row2Cells[2], "v");
    expect(c2v.text || getElementText(c2v)).toBe("0");
  });

  it("handles null and undefined values", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [[null, "hello", null]],
    };
    const styles = createStylesCollector();
    const ss = createSharedStrings();

    const xml = writeWorksheetXml(sheet, styles, ss).xml;
    const doc = parseXml(xml);

    const sheetData = findChild(doc, "sheetData");
    const rows = findChildren(sheetData, "row");
    expect(rows.length).toBe(1);

    const cells = findChildren(rows[0], "c");
    // Only the "hello" cell should be present (nulls are skipped without style)
    expect(cells.length).toBe(1);
    expect(cells[0].attrs["r"]).toBe("B1");
  });

  it("generates date cells with serial numbers", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [
        [new Date(Date.UTC(2024, 0, 15))], // Jan 15, 2024
      ],
    };
    const styles = createStylesCollector();
    const ss = createSharedStrings();

    const xml = writeWorksheetXml(sheet, styles, ss).xml;
    const doc = parseXml(xml);

    const sheetData = findChild(doc, "sheetData");
    const rows = findChildren(sheetData, "row");
    const cells = findChildren(rows[0], "c");

    // Date cell should be a number with style
    expect(cells[0].attrs["r"]).toBe("A1");
    expect(cells[0].attrs["t"]).toBeUndefined(); // dates are numbers
    expect(cells[0].attrs["s"]).toBeDefined(); // should have a style (date format)

    const val = findChild(cells[0], "v");
    const serial = parseFloat(val.text || getElementText(val));
    // Jan 15, 2024 = serial 45306 in 1900 date system
    expect(serial).toBe(45306);
  });

  it("writes formula cells", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [[10, 20]],
      cells: new Map([
        [
          "0,2",
          {
            formula: "A1+B1",
            formulaResult: 30,
          },
        ],
      ]),
    };
    const styles = createStylesCollector();
    const ss = createSharedStrings();

    const xml = writeWorksheetXml(sheet, styles, ss).xml;
    const doc = parseXml(xml);

    const sheetData = findChild(doc, "sheetData");
    const rows = findChildren(sheetData, "row");
    const cells = findChildren(rows[0], "c");

    expect(cells.length).toBe(3);

    // C1 should have formula
    const c1 = cells[2];
    expect(c1.attrs["r"]).toBe("C1");
    const fEl = findChild(c1, "f");
    expect(fEl).toBeDefined();
    expect(fEl.text || getElementText(fEl)).toBe("A1+B1");

    const vEl = findChild(c1, "v");
    expect(vEl.text || getElementText(vEl)).toBe("30");
  });

  it("writes merged cells", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Merged Header"]],
      merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 3 }],
    };
    const styles = createStylesCollector();
    const ss = createSharedStrings();

    const xml = writeWorksheetXml(sheet, styles, ss).xml;
    const doc = parseXml(xml);

    const mergeCellsEl = findChild(doc, "mergeCells");
    expect(mergeCellsEl).toBeDefined();
    expect(mergeCellsEl.attrs["count"]).toBe("1");

    const mergeCell = findChild(mergeCellsEl, "mergeCell");
    expect(mergeCell.attrs["ref"]).toBe("A1:D1");
  });

  it("writes freeze panes", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["Header"]],
      freezePane: { rows: 1, columns: 0 },
    };
    const styles = createStylesCollector();
    const ss = createSharedStrings();

    const xml = writeWorksheetXml(sheet, styles, ss).xml;
    const doc = parseXml(xml);

    const sheetViews = findChild(doc, "sheetViews");
    expect(sheetViews).toBeDefined();

    const sheetView = findChild(sheetViews, "sheetView");
    expect(sheetView).toBeDefined();

    const pane = findChild(sheetView, "pane");
    expect(pane).toBeDefined();
    expect(pane.attrs["ySplit"]).toBe("1");
    expect(pane.attrs["state"]).toBe("frozen");
    expect(pane.attrs["activePane"]).toBe("bottomLeft");
  });

  it("writes freeze panes for columns", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["A", "B"]],
      freezePane: { rows: 0, columns: 2 },
    };
    const styles = createStylesCollector();
    const ss = createSharedStrings();

    const xml = writeWorksheetXml(sheet, styles, ss).xml;
    const doc = parseXml(xml);

    const sheetViews = findChild(doc, "sheetViews");
    const sheetView = findChild(sheetViews, "sheetView");
    const pane = findChild(sheetView, "pane");
    expect(pane.attrs["xSplit"]).toBe("2");
    expect(pane.attrs["activePane"]).toBe("topRight");
  });

  it("writes freeze panes for both rows and columns", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["A"]],
      freezePane: { rows: 1, columns: 1 },
    };
    const styles = createStylesCollector();
    const ss = createSharedStrings();

    const xml = writeWorksheetXml(sheet, styles, ss).xml;
    const doc = parseXml(xml);

    const sheetViews = findChild(doc, "sheetViews");
    const sheetView = findChild(sheetViews, "sheetView");
    const pane = findChild(sheetView, "pane");
    expect(pane.attrs["xSplit"]).toBe("1");
    expect(pane.attrs["ySplit"]).toBe("1");
    expect(pane.attrs["activePane"]).toBe("bottomRight");
  });

  it("writes auto filter", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [
        ["Name", "Age"],
        ["Alice", 30],
      ],
      autoFilter: { range: "A1:B2" },
    };
    const styles = createStylesCollector();
    const ss = createSharedStrings();

    const xml = writeWorksheetXml(sheet, styles, ss).xml;
    const doc = parseXml(xml);

    const autoFilter = findChild(doc, "autoFilter");
    expect(autoFilter).toBeDefined();
    expect(autoFilter.attrs["ref"]).toBe("A1:B2");
  });

  it("writes column widths", () => {
    const sheet: WriteSheet = {
      name: "Test",
      columns: [{ width: 20 }, { width: 30 }, {}, { width: 15, hidden: true }],
      rows: [["A", "B", "C", "D"]],
    };
    const styles = createStylesCollector();
    const ss = createSharedStrings();

    const xml = writeWorksheetXml(sheet, styles, ss).xml;
    const doc = parseXml(xml);

    const colsEl = findChild(doc, "cols");
    expect(colsEl).toBeDefined();

    const colEls = findChildren(colsEl, "col");
    // columns with width or hidden: col[0], col[1], col[3]
    expect(colEls.length).toBe(3);

    expect(colEls[0].attrs["min"]).toBe("1");
    expect(colEls[0].attrs["max"]).toBe("1");
    expect(colEls[0].attrs["width"]).toBe("20");
    expect(colEls[0].attrs["customWidth"]).toBe("true");

    expect(colEls[2].attrs["min"]).toBe("4");
    expect(colEls[2].attrs["hidden"]).toBe("true");
  });

  it("writes empty sheet", () => {
    const sheet: WriteSheet = {
      name: "Empty",
      rows: [],
    };
    const styles = createStylesCollector();
    const ss = createSharedStrings();

    const xml = writeWorksheetXml(sheet, styles, ss).xml;
    const doc = parseXml(xml);

    const sheetData = findChild(doc, "sheetData");
    expect(sheetData).toBeDefined();
    // Should have no rows
    const rows = findChildren(sheetData, "row");
    expect(rows.length).toBe(0);
  });
});

// ── writeXlsx Full Integration ───────────────────────────────────────

describe("writeXlsx", () => {
  it("writes basic workbook with one sheet", async () => {
    const result = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            ["Name", "Age", "Active"],
            ["Alice", 30, true],
            ["Bob", 25, false],
          ],
        },
      ],
    });

    expect(result).toBeInstanceOf(Uint8Array);
    expect(result.length).toBeGreaterThan(0);

    // Verify it's a valid ZIP
    const zip = new ZipReader(result);
    const entries = zip.entries();
    expect(entries).toContain("[Content_Types].xml");
    expect(entries).toContain("_rels/.rels");
    expect(entries).toContain("xl/workbook.xml");
    expect(entries).toContain("xl/_rels/workbook.xml.rels");
    expect(entries).toContain("xl/styles.xml");
    expect(entries).toContain("xl/sharedStrings.xml");
    expect(entries).toContain("xl/worksheets/sheet1.xml");
  });

  it("verifies [Content_Types].xml structure", async () => {
    const result = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["hello"]] }],
    });

    const doc = await parseXmlFromZip(result, "[Content_Types].xml");
    expect(doc.local || doc.tag).toBe("Types");
  });

  it("verifies _rels/.rels structure", async () => {
    const result = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["hello"]] }],
    });

    const doc = await parseXmlFromZip(result, "_rels/.rels");
    const rels = findChildren(doc, "Relationship");
    expect(rels.length).toBe(3);
    expect(rels[0].attrs["Target"]).toBe("xl/workbook.xml");
    expect(rels[1].attrs["Target"]).toBe("docProps/core.xml");
    expect(rels[2].attrs["Target"]).toBe("docProps/app.xml");
  });

  it("verifies xl/workbook.xml has correct sheet names", async () => {
    const result = await writeXlsx({
      sheets: [
        { name: "Data", rows: [[1]] },
        { name: "Summary", rows: [[2]] },
      ],
    });

    const doc = await parseXmlFromZip(result, "xl/workbook.xml");
    const sheetsEl = findChild(doc, "sheets");
    const sheetEls = findChildren(sheetsEl, "sheet");
    expect(sheetEls.length).toBe(2);
    expect(sheetEls[0].attrs["name"]).toBe("Data");
    expect(sheetEls[1].attrs["name"]).toBe("Summary");
  });

  it("verifies worksheet has correct cell values", async () => {
    const result = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Hello", 42, true]],
        },
      ],
    });

    const doc = await parseXmlFromZip(result, "xl/worksheets/sheet1.xml");
    const sheetData = findChild(doc, "sheetData");
    const rows = findChildren(sheetData, "row");
    expect(rows.length).toBe(1);

    const cells = findChildren(rows[0], "c");
    expect(cells.length).toBe(3);

    // String cell
    expect(cells[0].attrs["t"]).toBe("s");
    // Number cell
    const numVal = findChild(cells[1], "v");
    expect(numVal.text || getElementText(numVal)).toBe("42");
    // Boolean cell
    expect(cells[2].attrs["t"]).toBe("b");
  });

  it("deduplicates shared strings", async () => {
    const result = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            ["hello", "world", "hello"],
            ["world", "hello", "new"],
          ],
        },
      ],
    });

    const doc = await parseXmlFromZip(result, "xl/sharedStrings.xml");
    const siElements = findChildren(doc, "si");
    // "hello", "world", "new" = 3 unique strings
    expect(siElements.length).toBe(3);

    // Verify indices in worksheet
    const wsDoc = await parseXmlFromZip(result, "xl/worksheets/sheet1.xml");
    const sheetData = findChild(wsDoc, "sheetData");
    const rows = findChildren(sheetData, "row");
    const row1Cells = findChildren(rows[0], "c");

    // A1 and C1 should have same shared string index
    const a1v = findChild(row1Cells[0], "v");
    const c1v = findChild(row1Cells[2], "v");
    expect(a1v.text || getElementText(a1v)).toBe(c1v.text || getElementText(c1v));
  });

  it("writes dates with serial number and numFmt style", async () => {
    const date = new Date(Date.UTC(2024, 0, 15)); // Jan 15, 2024
    const result = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [[date]],
        },
      ],
    });

    // Check styles.xml has a date numFmt
    const stylesDoc = await parseXmlFromZip(result, "xl/styles.xml");
    const numFmtsEl = findChild(stylesDoc, "numFmts");
    expect(numFmtsEl).toBeDefined();

    const numFmt = findChild(numFmtsEl, "numFmt");
    expect(numFmt.attrs["formatCode"]).toBe("yyyy-mm-dd");
    expect(Number(numFmt.attrs["numFmtId"])).toBeGreaterThanOrEqual(164);

    // Check worksheet cell has serial value and style reference
    const wsDoc = await parseXmlFromZip(result, "xl/worksheets/sheet1.xml");
    const sheetData = findChild(wsDoc, "sheetData");
    const rows = findChildren(sheetData, "row");
    const cells = findChildren(rows[0], "c");
    expect(cells[0].attrs["s"]).toBeDefined(); // has style
    expect(cells[0].attrs["t"]).toBeUndefined(); // not string type
  });

  it("writes with cell styles (bold, colors, borders)", async () => {
    const result = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Styled"]],
          cells: new Map([
            [
              "0,0",
              {
                value: "Styled",
                style: {
                  font: { bold: true, color: { rgb: "FF0000" } },
                  fill: {
                    type: "pattern",
                    pattern: "solid",
                    fgColor: { rgb: "FFFF00" },
                  },
                  border: {
                    top: { style: "thin" },
                    bottom: { style: "thin" },
                    left: { style: "thin" },
                    right: { style: "thin" },
                  },
                },
              },
            ],
          ]),
        },
      ],
    });

    const stylesDoc = await parseXmlFromZip(result, "xl/styles.xml");

    // Check custom font exists
    const fontsEl = findChild(stylesDoc, "fonts");
    expect(Number(fontsEl.attrs["count"])).toBeGreaterThan(1);

    // Check custom fill exists
    const fillsEl = findChild(stylesDoc, "fills");
    expect(Number(fillsEl.attrs["count"])).toBeGreaterThan(2);

    // Check custom border exists
    const bordersEl = findChild(stylesDoc, "borders");
    expect(Number(bordersEl.attrs["count"])).toBeGreaterThan(1);

    // Check cellXfs has the styled entry
    const cellXfsEl = findChild(stylesDoc, "cellXfs");
    expect(Number(cellXfsEl.attrs["count"])).toBeGreaterThan(1);
  });

  it("writes column widths", async () => {
    const result = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          columns: [{ width: 20 }, { width: 35 }],
          rows: [["A", "B"]],
        },
      ],
    });

    const wsDoc = await parseXmlFromZip(result, "xl/worksheets/sheet1.xml");
    const colsEl = findChild(wsDoc, "cols");
    expect(colsEl).toBeDefined();

    const colEls = findChildren(colsEl, "col");
    expect(colEls.length).toBe(2);
    expect(colEls[0].attrs["width"]).toBe("20");
    expect(colEls[1].attrs["width"]).toBe("35");
  });

  it("writes merged cells", async () => {
    const result = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            ["Header", null, null],
            ["A", "B", "C"],
          ],
          merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }],
        },
      ],
    });

    const wsDoc = await parseXmlFromZip(result, "xl/worksheets/sheet1.xml");
    const mergeCellsEl = findChild(wsDoc, "mergeCells");
    expect(mergeCellsEl).toBeDefined();
    expect(mergeCellsEl.attrs["count"]).toBe("1");

    const mergeCell = findChild(mergeCellsEl, "mergeCell");
    expect(mergeCell.attrs["ref"]).toBe("A1:C1");
  });

  it("writes freeze panes", async () => {
    const result = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Header"], ["Data"]],
          freezePane: { rows: 1 },
        },
      ],
    });

    const wsDoc = await parseXmlFromZip(result, "xl/worksheets/sheet1.xml");
    const sheetViews = findChild(wsDoc, "sheetViews");
    const sheetView = findChild(sheetViews, "sheetView");
    const pane = findChild(sheetView, "pane");
    expect(pane).toBeDefined();
    expect(pane.attrs["ySplit"]).toBe("1");
    expect(pane.attrs["state"]).toBe("frozen");
  });

  it("writes auto filter", async () => {
    const result = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            ["Name", "Value"],
            ["A", 1],
            ["B", 2],
          ],
          autoFilter: { range: "A1:B3" },
        },
      ],
    });

    const wsDoc = await parseXmlFromZip(result, "xl/worksheets/sheet1.xml");
    const autoFilter = findChild(wsDoc, "autoFilter");
    expect(autoFilter).toBeDefined();
    expect(autoFilter.attrs["ref"]).toBe("A1:B3");
  });

  it("writes multiple sheets", async () => {
    const result = await writeXlsx({
      sheets: [
        { name: "Sheet1", rows: [["A"]] },
        { name: "Sheet2", rows: [["B"]] },
        { name: "Sheet3", rows: [["C"]] },
      ],
    });

    const zip = new ZipReader(result);
    expect(zip.has("xl/worksheets/sheet1.xml")).toBe(true);
    expect(zip.has("xl/worksheets/sheet2.xml")).toBe(true);
    expect(zip.has("xl/worksheets/sheet3.xml")).toBe(true);

    // Verify workbook lists all sheets
    const wbDoc = await parseXmlFromZip(result, "xl/workbook.xml");
    const sheetsEl = findChild(wbDoc, "sheets");
    const sheetEls = findChildren(sheetsEl, "sheet");
    expect(sheetEls.length).toBe(3);
    expect(sheetEls[0].attrs["name"]).toBe("Sheet1");
    expect(sheetEls[1].attrs["name"]).toBe("Sheet2");
    expect(sheetEls[2].attrs["name"]).toBe("Sheet3");

    // Verify relationships
    const relsDoc = await parseXmlFromZip(result, "xl/_rels/workbook.xml.rels");
    const rels = findChildren(relsDoc, "Relationship");
    // 3 sheets + styles + shared strings + theme = 6
    expect(rels.length).toBe(6);
  });

  it("writes empty sheet", async () => {
    const result = await writeXlsx({
      sheets: [{ name: "Empty", rows: [] }],
    });

    const wsDoc = await parseXmlFromZip(result, "xl/worksheets/sheet1.xml");
    const sheetData = findChild(wsDoc, "sheetData");
    expect(sheetData).toBeDefined();
  });

  it("writes from object data with column keys", async () => {
    const result = await writeXlsx({
      sheets: [
        {
          name: "People",
          columns: [
            { key: "name", header: "Name", width: 20 },
            { key: "age", header: "Age", width: 10 },
          ],
          data: [
            { name: "Alice", age: 30 },
            { name: "Bob", age: 25 },
          ],
        },
      ],
    });

    // Verify shared strings contain headers and names
    const ssDoc = await parseXmlFromZip(result, "xl/sharedStrings.xml");
    const siElements = findChildren(ssDoc, "si");
    const stringValues = siElements.map((si: any) => {
      const t = findChild(si, "t");
      return t.text || getElementText(t);
    });
    expect(stringValues).toContain("Name");
    expect(stringValues).toContain("Age");
    expect(stringValues).toContain("Alice");
    expect(stringValues).toContain("Bob");

    // Verify worksheet has 3 rows (header + 2 data rows)
    const wsDoc = await parseXmlFromZip(result, "xl/worksheets/sheet1.xml");
    const sheetData = findChild(wsDoc, "sheetData");
    const rows = findChildren(sheetData, "row");
    expect(rows.length).toBe(3);
  });

  it("writes formula cells", async () => {
    const result = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [[10, 20, null]],
          cells: new Map([
            [
              "0,2",
              {
                formula: "SUM(A1:B1)",
                formulaResult: 30,
              },
            ],
          ]),
        },
      ],
    });

    const wsDoc = await parseXmlFromZip(result, "xl/worksheets/sheet1.xml");
    const sheetData = findChild(wsDoc, "sheetData");
    const rows = findChildren(sheetData, "row");
    const cells = findChildren(rows[0], "c");

    // Find C1 cell
    const c1 = cells.find((c: any) => c.attrs["r"] === "C1");
    expect(c1).toBeDefined();

    const fEl = findChild(c1, "f");
    expect(fEl).toBeDefined();
    expect(fEl.text || getElementText(fEl)).toBe("SUM(A1:B1)");

    const vEl = findChild(c1, "v");
    expect(vEl.text || getElementText(vEl)).toBe("30");
  });

  it("writes large sheet (1000 rows)", async () => {
    const rows: CellValue[][] = [];
    for (let i = 0; i < 1000; i++) {
      rows.push([`Row ${i}`, i, i % 2 === 0]);
    }

    const result = await writeXlsx({
      sheets: [{ name: "Large", rows }],
    });

    expect(result).toBeInstanceOf(Uint8Array);
    expect(result.length).toBeGreaterThan(0);

    // Verify it's a valid ZIP and has the worksheet
    const zip = new ZipReader(result);
    expect(zip.has("xl/worksheets/sheet1.xml")).toBe(true);

    // Spot check first and last row
    const wsDoc = await parseXmlFromZip(result, "xl/worksheets/sheet1.xml");
    const sheetData = findChild(wsDoc, "sheetData");
    const rowEls = findChildren(sheetData, "row");
    expect(rowEls.length).toBe(1000);
    expect(rowEls[0].attrs["r"]).toBe("1");
    expect(rowEls[999].attrs["r"]).toBe("1000");
  });

  it("output is valid ZIP archive", async () => {
    const result = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    });

    // Should start with PK signature
    expect(result[0]).toBe(0x50); // P
    expect(result[1]).toBe(0x4b); // K

    // Should be readable by ZipReader
    const zip = new ZipReader(result);
    const entries = zip.entries();
    expect(entries.length).toBeGreaterThan(0);
  });

  it("writes with custom number formats", async () => {
    const result = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [[1234.56]],
          cells: new Map([
            [
              "0,0",
              {
                value: 1234.56,
                style: { numFmt: "#,##0.00" },
              },
            ],
          ]),
        },
      ],
    });

    const stylesDoc = await parseXmlFromZip(result, "xl/styles.xml");
    const numFmtsEl = findChild(stylesDoc, "numFmts");
    expect(numFmtsEl).toBeDefined();

    const numFmt = findChild(numFmtsEl, "numFmt");
    expect(numFmt.attrs["formatCode"]).toBe("#,##0.00");
    expect(Number(numFmt.attrs["numFmtId"])).toBeGreaterThanOrEqual(164);
  });

  it("writes hidden sheets", async () => {
    const result = await writeXlsx({
      sheets: [
        { name: "Visible", rows: [["A"]] },
        { name: "Hidden", rows: [["B"]], hidden: true },
      ],
    });

    const wbDoc = await parseXmlFromZip(result, "xl/workbook.xml");
    const sheetsEl = findChild(wbDoc, "sheets");
    const sheetEls = findChildren(sheetsEl, "sheet");
    expect(sheetEls[0].attrs["state"]).toBeUndefined();
    expect(sheetEls[1].attrs["state"]).toBe("hidden");
  });

  it("style deduplication - same style reuses xf index", async () => {
    const boldStyle: CellStyle = { font: { bold: true } };

    const result = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["A", "B"]],
          cells: new Map([
            ["0,0", { value: "A", style: boldStyle }],
            ["0,1", { value: "B", style: boldStyle }],
          ]),
        },
      ],
    });

    // Verify styles.xml has exactly 2 xfs (default + bold)
    const stylesDoc = await parseXmlFromZip(result, "xl/styles.xml");
    const cellXfsEl = findChild(stylesDoc, "cellXfs");
    expect(cellXfsEl.attrs["count"]).toBe("2");

    // Verify both cells reference the same style index
    const wsDoc = await parseXmlFromZip(result, "xl/worksheets/sheet1.xml");
    const sheetData = findChild(wsDoc, "sheetData");
    const rows = findChildren(sheetData, "row");
    const cells = findChildren(rows[0], "c");

    expect(cells[0].attrs["s"]).toBe(cells[1].attrs["s"]);
  });

  it("writes with default font", async () => {
    const result = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
      defaultFont: { name: "Arial", size: 12 },
    });

    const stylesDoc = await parseXmlFromZip(result, "xl/styles.xml");
    const fontsEl = findChild(stylesDoc, "fonts");
    const fontEls = findChildren(fontsEl, "font");

    // Default font should be Arial 12
    const nameEl = findChild(fontEls[0], "name");
    expect(nameEl.attrs["val"]).toBe("Arial");
    const sizeEl = findChild(fontEls[0], "sz");
    expect(sizeEl.attrs["val"]).toBe("12");
  });

  it("handles special characters in strings", async () => {
    const result = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [['He said "hello" & goodbye'], ["<tag>content</tag>"]],
        },
      ],
    });

    // Should be valid ZIP (no XML corruption)
    const zip = new ZipReader(result);
    expect(zip.has("xl/sharedStrings.xml")).toBe(true);

    // Parse should succeed (XML escaping worked)
    const ssDoc = await parseXmlFromZip(result, "xl/sharedStrings.xml");
    const siElements = findChildren(ssDoc, "si");
    expect(siElements.length).toBe(2);

    const t1 = findChild(siElements[0], "t");
    expect(t1.text || getElementText(t1)).toBe('He said "hello" & goodbye');

    const t2 = findChild(siElements[1], "t");
    expect(t2.text || getElementText(t2)).toBe("<tag>content</tag>");
  });

  it("omits sharedStrings.xml when no strings", async () => {
    const result = await writeXlsx({
      sheets: [
        {
          name: "Numbers",
          rows: [
            [1, 2, 3],
            [4, 5, 6],
          ],
        },
      ],
    });

    const zip = new ZipReader(result);
    expect(zip.has("xl/sharedStrings.xml")).toBe(false);

    // Content types should not have shared strings override
    const ctDoc = await parseXmlFromZip(result, "[Content_Types].xml");
    const overrides = findChildren(ctDoc, "Override");
    const ssOverride = overrides.find((o: any) => o.attrs["PartName"] === "/xl/sharedStrings.xml");
    expect(ssOverride).toBeUndefined();
  });

  it("writes alignment styles", async () => {
    const result = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Centered"]],
          cells: new Map([
            [
              "0,0",
              {
                value: "Centered",
                style: {
                  alignment: {
                    horizontal: "center",
                    vertical: "center",
                    wrapText: true,
                  },
                },
              },
            ],
          ]),
        },
      ],
    });

    const stylesDoc = await parseXmlFromZip(result, "xl/styles.xml");
    const cellXfsEl = findChild(stylesDoc, "cellXfs");
    const xfEls = findChildren(cellXfsEl, "xf");

    // Find the xf with alignment
    const alignedXf = xfEls.find((xf: any) => xf.attrs["applyAlignment"] === "true");
    expect(alignedXf).toBeDefined();

    const alignmentEl = findChild(alignedXf, "alignment");
    expect(alignmentEl).toBeDefined();
    expect(alignmentEl.attrs["horizontal"]).toBe("center");
    expect(alignmentEl.attrs["vertical"]).toBe("center");
    expect(alignmentEl.attrs["wrapText"]).toBe("true");
  });

  it("writes protection styles", async () => {
    const result = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Protected"]],
          cells: new Map([
            [
              "0,0",
              {
                value: "Protected",
                style: {
                  protection: { locked: true, hidden: true },
                },
              },
            ],
          ]),
        },
      ],
    });

    const stylesDoc = await parseXmlFromZip(result, "xl/styles.xml");
    const cellXfsEl = findChild(stylesDoc, "cellXfs");
    const xfEls = findChildren(cellXfsEl, "xf");

    const protectedXf = xfEls.find((xf: any) => xf.attrs["applyProtection"] === "true");
    expect(protectedXf).toBeDefined();

    const protectionEl = findChild(protectedXf, "protection");
    expect(protectionEl).toBeDefined();
    expect(protectionEl.attrs["locked"]).toBe("1");
    expect(protectionEl.attrs["hidden"]).toBe("1");
  });

  it("writes gradient fill", async () => {
    const result = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Gradient"]],
          cells: new Map([
            [
              "0,0",
              {
                value: "Gradient",
                style: {
                  fill: {
                    type: "gradient",
                    degree: 90,
                    stops: [
                      { position: 0, color: { rgb: "FFFFFF" } },
                      { position: 1, color: { rgb: "000000" } },
                    ],
                  },
                },
              },
            ],
          ]),
        },
      ],
    });

    const stylesDoc = await parseXmlFromZip(result, "xl/styles.xml");
    const fillsEl = findChild(stylesDoc, "fills");
    const fillEls = findChildren(fillsEl, "fill");
    // 2 defaults + 1 gradient = 3
    expect(fillEls.length).toBe(3);

    const gradientFill = findChild(fillEls[2], "gradientFill");
    expect(gradientFill).toBeDefined();
    expect(gradientFill.attrs["degree"]).toBe("90");
  });

  it("writes multiple number format styles", async () => {
    const result = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [[1234.5, 0.75, 42]],
          cells: new Map([
            ["0,0", { value: 1234.5, style: { numFmt: "#,##0.00" } }],
            ["0,1", { value: 0.75, style: { numFmt: "0.00%" } }],
            ["0,2", { value: 42, style: { numFmt: "#,##0.00" } }],
          ]),
        },
      ],
    });

    const stylesDoc = await parseXmlFromZip(result, "xl/styles.xml");
    const numFmtsEl = findChild(stylesDoc, "numFmts");
    const numFmtEls = findChildren(numFmtsEl, "numFmt");

    // 2 unique formats: #,##0.00 and 0.00%
    expect(numFmtEls.length).toBe(2);
    expect(numFmtEls[0].attrs["numFmtId"]).toBe("164");
    expect(numFmtEls[1].attrs["numFmtId"]).toBe("165");
  });

  it("writes mixed data types in same row", async () => {
    const date = new Date(Date.UTC(2024, 5, 15));
    const result = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["text", 42, true, date, null, false]],
        },
      ],
    });

    const wsDoc = await parseXmlFromZip(result, "xl/worksheets/sheet1.xml");
    const sheetData = findChild(wsDoc, "sheetData");
    const rows = findChildren(sheetData, "row");
    const cells = findChildren(rows[0], "c");

    // A1=string, B1=number, C1=bool, D1=date, E1=null(skip), F1=bool
    // null is skipped, so 5 cells
    expect(cells.length).toBe(5);

    expect(cells[0].attrs["t"]).toBe("s"); // string
    expect(cells[1].attrs["t"]).toBeUndefined(); // number (no type attr)
    expect(cells[2].attrs["t"]).toBe("b"); // boolean
    expect(cells[3].attrs["t"]).toBeUndefined(); // date as number
    expect(cells[3].attrs["s"]).toBeDefined(); // date has style
    expect(cells[4].attrs["t"]).toBe("b"); // boolean
    expect(cells[4].attrs["r"]).toBe("F1"); // column F (E is skipped)
  });

  it("handles cell overrides on top of row data", async () => {
    const result = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Original", 100]],
          cells: new Map([["0,0", { value: "Overridden", style: { font: { bold: true } } }]]),
        },
      ],
    });

    const wsDoc = await parseXmlFromZip(result, "xl/worksheets/sheet1.xml");
    const sheetData = findChild(wsDoc, "sheetData");
    const rows = findChildren(sheetData, "row");
    const cells = findChildren(rows[0], "c");

    // A1 should be "Overridden" with bold style
    expect(cells[0].attrs["t"]).toBe("s");
    expect(cells[0].attrs["s"]).toBeDefined(); // has style

    // Verify shared string is "Overridden"
    const ssDoc = await parseXmlFromZip(result, "xl/sharedStrings.xml");
    const siElements = findChildren(ssDoc, "si");
    const v = findChild(cells[0], "v");
    const idx = parseInt(v.text || getElementText(v), 10);
    const t = findChild(siElements[idx], "t");
    expect(t.text || getElementText(t)).toBe("Overridden");
  });

  it("writes underline style variations", async () => {
    const result = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Single", "Double"]],
          cells: new Map([
            ["0,0", { value: "Single", style: { font: { underline: true } } }],
            ["0,1", { value: "Double", style: { font: { underline: "double" } } }],
          ]),
        },
      ],
    });

    const stylesDoc = await parseXmlFromZip(result, "xl/styles.xml");
    const fontsEl = findChild(stylesDoc, "fonts");
    const fontEls = findChildren(fontsEl, "font");

    // default + single underline + double underline = 3
    expect(fontEls.length).toBe(3);

    // Single underline: <u/> (no val attribute)
    const singleFont = fontEls[1];
    const singleU = findChild(singleFont, "u");
    expect(singleU).toBeDefined();
    expect(singleU.attrs["val"]).toBeUndefined();

    // Double underline: <u val="double"/>
    const doubleFont = fontEls[2];
    const doubleU = findChild(doubleFont, "u");
    expect(doubleU).toBeDefined();
    expect(doubleU.attrs["val"]).toBe("double");
  });
});
