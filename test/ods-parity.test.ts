import { describe, it, expect } from "vitest";
import { writeOds } from "../src/ods/writer";
import { readOds } from "../src/ods/reader";
import { writeXlsx } from "../src/xlsx/writer";
import { readXlsx } from "../src/xlsx/reader";
import { ZipReader } from "../src/zip/reader";
import { parseXml } from "../src/xml/parser";
import type { WriteSheet, Cell, MergeRange } from "../src/_types";

// ── Helpers ──────────────────────────────────────────────────────────

const decoder = new TextDecoder("utf-8");

async function extractFile(data: Uint8Array, path: string): Promise<string> {
  const zip = new ZipReader(data);
  const raw = await zip.extract(path);
  return decoder.decode(raw);
}

async function parseXmlFromZip(data: Uint8Array, path: string) {
  const xml = await extractFile(data, path);
  return parseXml(xml);
}

function findChild(el: { children: Array<unknown> }, localName: string): any {
  return el.children.find((c: any) => typeof c !== "string" && (c.local || c.tag) === localName);
}

function findChildren(el: { children: Array<unknown> }, localName: string): any[] {
  return el.children.filter((c: any) => typeof c !== "string" && (c.local || c.tag) === localName);
}

// ── Merged Cells ────────────────────────────────────────────────────

describe("ODS parity — merged cells", () => {
  it("write ODS with merged cells → read back → merges preserved", async () => {
    const merges: MergeRange[] = [
      { startRow: 0, startCol: 0, endRow: 0, endCol: 2 }, // A1:C1
      { startRow: 1, startCol: 1, endRow: 2, endCol: 2 }, // B2:C3
    ];

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        ["Merged Header", null, null, "D1"],
        ["A2", "Merged Block", null],
        ["A3", null, null],
      ],
      merges,
    };

    const data = await writeOds({ sheets: [sheet] });
    const wb = await readOds(data);

    expect(wb.sheets[0].merges).toBeDefined();
    expect(wb.sheets[0].merges).toHaveLength(2);

    // First merge: A1:C1
    const m1 = wb.sheets[0].merges!.find((m) => m.startRow === 0 && m.startCol === 0);
    expect(m1).toBeDefined();
    expect(m1!.endRow).toBe(0);
    expect(m1!.endCol).toBe(2);

    // Second merge: B2:C3
    const m2 = wb.sheets[0].merges!.find((m) => m.startRow === 1 && m.startCol === 1);
    expect(m2).toBeDefined();
    expect(m2!.endRow).toBe(2);
    expect(m2!.endCol).toBe(2);
  });

  it("writes number-columns-spanned and number-rows-spanned attributes", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["A", null, null]],
      merges: [{ startRow: 0, startCol: 0, endRow: 1, endCol: 2 }],
    };

    const data = await writeOds({ sheets: [sheet] });
    const contentXml = await extractFile(data, "content.xml");

    expect(contentXml).toContain('table:number-columns-spanned="3"');
    expect(contentXml).toContain('table:number-rows-spanned="2"');
  });

  it("writes covered-table-cell for cells covered by a merge", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        ["Merged", null, null],
        [null, null, null],
      ],
      merges: [{ startRow: 0, startCol: 0, endRow: 1, endCol: 2 }],
    };

    const data = await writeOds({ sheets: [sheet] });
    const contentXml = await extractFile(data, "content.xml");

    expect(contentXml).toContain("table:covered-table-cell");
  });

  it("reader handles covered-table-cell correctly (null values)", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Merged", null, null, "D1"]],
      merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }],
    };

    const data = await writeOds({ sheets: [sheet] });
    const wb = await readOds(data);

    // The merged cell value should be preserved, covered cells are null
    expect(wb.sheets[0].rows[0][0]).toBe("Merged");
    expect(wb.sheets[0].rows[0][1]).toBe(null);
    expect(wb.sheets[0].rows[0][2]).toBe(null);
    expect(wb.sheets[0].rows[0][3]).toBe("D1");
  });

  it("merge spanning multiple rows works correctly", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        ["Top", "Data"],
        [null, "More"],
        [null, "Even more"],
      ],
      merges: [{ startRow: 0, startCol: 0, endRow: 2, endCol: 0 }],
    };

    const data = await writeOds({ sheets: [sheet] });
    const wb = await readOds(data);

    expect(wb.sheets[0].merges).toHaveLength(1);
    expect(wb.sheets[0].merges![0]).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 2,
      endCol: 0,
    });

    expect(wb.sheets[0].rows[0][0]).toBe("Top");
    expect(wb.sheets[0].rows[0][1]).toBe("Data");
    expect(wb.sheets[0].rows[1][1]).toBe("More");
  });
});

// ── Hyperlinks ──────────────────────────────────────────────────────

describe("ODS parity — hyperlinks", () => {
  it("write ODS with hyperlinks → read back → hyperlinks preserved", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: "Click here",
      hyperlink: { target: "https://example.com", display: "Click here" },
    });
    cells.set("1,0", {
      value: "Google",
      hyperlink: { target: "https://google.com", display: "Google" },
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Click here"], ["Google"]],
      cells,
    };

    const data = await writeOds({ sheets: [sheet] });
    const wb = await readOds(data);

    expect(wb.sheets[0].cells).toBeDefined();

    const cell00 = wb.sheets[0].cells!.get("0,0");
    expect(cell00).toBeDefined();
    expect(cell00!.hyperlink).toBeDefined();
    expect(cell00!.hyperlink!.target).toBe("https://example.com");

    const cell10 = wb.sheets[0].cells!.get("1,0");
    expect(cell10).toBeDefined();
    expect(cell10!.hyperlink).toBeDefined();
    expect(cell10!.hyperlink!.target).toBe("https://google.com");
  });

  it("reader extracts text:a hyperlinks", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: "Link Text",
      hyperlink: { target: "https://test.org", display: "Link Text" },
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Link Text"]],
      cells,
    };

    const data = await writeOds({ sheets: [sheet] });

    // Verify the XML has text:a element
    const contentXml = await extractFile(data, "content.xml");
    expect(contentXml).toContain("text:a");
    expect(contentXml).toContain("https://test.org");

    // Verify reader extracts hyperlink
    const wb = await readOds(data);
    const cell = wb.sheets[0].cells!.get("0,0");
    expect(cell).toBeDefined();
    expect(cell!.hyperlink!.target).toBe("https://test.org");
    expect(cell!.hyperlink!.display).toBe("Link Text");
  });

  it("hyperlink cells still have correct text values in rows", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: "Visit",
      hyperlink: { target: "https://example.com" },
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Visit", "Normal"]],
      cells,
    };

    const data = await writeOds({ sheets: [sheet] });
    const wb = await readOds(data);

    expect(wb.sheets[0].rows[0][0]).toBe("Visit");
    expect(wb.sheets[0].rows[0][1]).toBe("Normal");
  });
});

// ── Basic Styles ────────────────────────────────────────────────────

describe("ODS parity — styles", () => {
  it("write ODS with bold/italic/color styles → read back with readStyles", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: "Bold",
      style: { font: { bold: true } },
    });
    cells.set("0,1", {
      value: "Italic",
      style: { font: { italic: true } },
    });
    cells.set("0,2", {
      value: "Red",
      style: { font: { color: { rgb: "FF0000" } } },
    });
    cells.set("1,0", {
      value: "Big",
      style: { font: { size: 24 } },
    });
    cells.set("1,1", {
      value: "BG",
      style: {
        fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "00FF00" } },
      },
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        ["Bold", "Italic", "Red"],
        ["Big", "BG"],
      ],
      cells,
    };

    const data = await writeOds({ sheets: [sheet] });

    // Verify style elements in content.xml
    const contentXml = await extractFile(data, "content.xml");
    expect(contentXml).toContain("fo:font-weight");
    expect(contentXml).toContain("fo:font-style");
    expect(contentXml).toContain("fo:color");
    expect(contentXml).toContain("fo:font-size");
    expect(contentXml).toContain("fo:background-color");

    // Read back with readStyles
    const wb = await readOds(data, { readStyles: true });

    const cell00 = wb.sheets[0].cells!.get("0,0");
    expect(cell00).toBeDefined();
    expect(cell00!.style?.font?.bold).toBe(true);

    const cell01 = wb.sheets[0].cells!.get("0,1");
    expect(cell01).toBeDefined();
    expect(cell01!.style?.font?.italic).toBe(true);

    const cell02 = wb.sheets[0].cells!.get("0,2");
    expect(cell02).toBeDefined();
    expect(cell02!.style?.font?.color?.rgb).toBe("FF0000");

    const cell10 = wb.sheets[0].cells!.get("1,0");
    expect(cell10).toBeDefined();
    expect(cell10!.style?.font?.size).toBe(24);

    const cell11 = wb.sheets[0].cells!.get("1,1");
    expect(cell11).toBeDefined();
    expect(cell11!.style?.fill?.type).toBe("pattern");
    if (cell11!.style?.fill?.type === "pattern") {
      expect(cell11!.style.fill.fgColor?.rgb).toBe("00FF00");
    }
  });

  it("styles are not collected when readStyles is false (default)", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: "Bold",
      style: { font: { bold: true } },
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Bold"]],
      cells,
    };

    const data = await writeOds({ sheets: [sheet] });
    const wb = await readOds(data); // default: readStyles = false

    // Without readStyles, cells map should not contain style-only entries
    // (the cell has no formula or hyperlink, so it's only interesting for styles)
    if (wb.sheets[0].cells) {
      const cell00 = wb.sheets[0].cells.get("0,0");
      if (cell00) {
        expect(cell00.style).toBeUndefined();
      }
    }
  });

  it("generates unique style names for different styles", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: "Bold",
      style: { font: { bold: true } },
    });
    cells.set("0,1", {
      value: "Italic",
      style: { font: { italic: true } },
    });
    // Same style as 0,0 — should reuse the style name
    cells.set("1,0", {
      value: "Also Bold",
      style: { font: { bold: true } },
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Bold", "Italic"], ["Also Bold"]],
      cells,
    };

    const data = await writeOds({ sheets: [sheet] });
    const contentDoc = await parseXmlFromZip(data, "content.xml");
    const autoStyles = findChild(contentDoc, "automatic-styles");
    const styles = findChildren(autoStyles, "style");

    // Should have exactly 2 styles (bold and italic), not 3
    expect(styles.length).toBe(2);
  });

  it("style references cells via table:style-name", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: "Styled",
      style: { font: { bold: true } },
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Styled"]],
      cells,
    };

    const data = await writeOds({ sheets: [sheet] });
    const contentXml = await extractFile(data, "content.xml");

    // The cell should reference the style
    expect(contentXml).toContain("table:style-name=");
  });
});

// ── Formulas ────────────────────────────────────────────────────────

describe("ODS parity — formulas", () => {
  it("write ODS with formulas → read back → formula text preserved", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("1,0", {
      value: 55,
      formula: "SUM(A1:A1)",
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [
        [10],
        [55], // formula result
      ],
      cells,
    };

    const data = await writeOds({ sheets: [sheet] });
    const wb = await readOds(data);

    expect(wb.sheets[0].cells).toBeDefined();
    const cell = wb.sheets[0].cells!.get("1,0");
    expect(cell).toBeDefined();
    expect(cell!.formula).toBe("SUM(A1:A1)");
  });

  it("writes table:formula attribute with of:= prefix", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: 42,
      formula: "SUM(B1:B10)",
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [[42]],
      cells,
    };

    const data = await writeOds({ sheets: [sheet] });
    const contentXml = await extractFile(data, "content.xml");

    expect(contentXml).toContain("table:formula=");
    expect(contentXml).toContain("of:=SUM([.B1:.B10])");
  });

  it("converts ODS cell references back to Excel style on read", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,2", {
      value: 100,
      formula: "A1+B1",
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [[50, 50, 100]],
      cells,
    };

    const data = await writeOds({ sheets: [sheet] });

    // Verify ODS format is written
    const contentXml = await extractFile(data, "content.xml");
    expect(contentXml).toContain("of:=[.A1]+[.B1]");

    // Read back and verify Excel format
    const wb = await readOds(data);
    const cell = wb.sheets[0].cells!.get("0,2");
    expect(cell!.formula).toBe("A1+B1");
  });

  it("handles range formulas like SUM(A1:A10)", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: 0,
      formula: "SUM(A2:A10)",
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [[0]],
      cells,
    };

    const data = await writeOds({ sheets: [sheet] });
    const wb = await readOds(data);

    const cell = wb.sheets[0].cells!.get("0,0");
    expect(cell!.formula).toBe("SUM(A2:A10)");
  });
});

// ── Cross-format Parity ─────────────────────────────────────────────

describe("ODS parity — cross-format", () => {
  it("write XLSX with merges → read, write same data as ODS → read, compare", async () => {
    const merges: MergeRange[] = [
      { startRow: 0, startCol: 0, endRow: 0, endCol: 2 },
      { startRow: 1, startCol: 0, endRow: 2, endCol: 1 },
    ];

    const rows = [
      ["Header", null, null, "D1"],
      ["Block", null, "C2"],
      [null, null, "C3"],
    ];

    // Write as XLSX
    const xlsxData = await writeXlsx({
      sheets: [{ name: "Sheet1", rows, merges }],
    });
    const xlsxWb = await readXlsx(xlsxData);

    // Write as ODS
    const odsData = await writeOds({
      sheets: [{ name: "Sheet1", rows, merges }],
    });
    const odsWb = await readOds(odsData);

    // Compare merges
    expect(odsWb.sheets[0].merges).toBeDefined();
    expect(odsWb.sheets[0].merges).toHaveLength(xlsxWb.sheets[0].merges!.length);

    // Sort for comparison
    const sortMerge = (a: MergeRange, b: MergeRange) =>
      a.startRow - b.startRow || a.startCol - b.startCol;
    const xlsxMerges = [...xlsxWb.sheets[0].merges!].sort(sortMerge);
    const odsMerges = [...odsWb.sheets[0].merges!].sort(sortMerge);

    for (let i = 0; i < xlsxMerges.length; i++) {
      expect(odsMerges[i].startRow).toBe(xlsxMerges[i].startRow);
      expect(odsMerges[i].startCol).toBe(xlsxMerges[i].startCol);
      expect(odsMerges[i].endRow).toBe(xlsxMerges[i].endRow);
      expect(odsMerges[i].endCol).toBe(xlsxMerges[i].endCol);
    }

    // Compare data values (non-null cells)
    expect(odsWb.sheets[0].rows[0][0]).toBe(xlsxWb.sheets[0].rows[0][0]);
    expect(odsWb.sheets[0].rows[0][3]).toBe(xlsxWb.sheets[0].rows[0][3]);
  });

  it("XLSX and ODS produce same row/cell values for basic data", async () => {
    const rows = [
      ["Name", "Age", "Active"],
      ["Alice", 30, true],
      ["Bob", 25, false],
    ];

    const xlsxData = await writeXlsx({ sheets: [{ name: "Data", rows }] });
    const xlsxWb = await readXlsx(xlsxData);

    const odsData = await writeOds({ sheets: [{ name: "Data", rows }] });
    const odsWb = await readOds(odsData);

    // Compare all rows
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < rows[r].length; c++) {
        expect(odsWb.sheets[0].rows[r][c]).toEqual(xlsxWb.sheets[0].rows[r][c]);
      }
    }
  });
});

// ── Combined Features ───────────────────────────────────────────────

describe("ODS parity — combined features", () => {
  it("sheet with merges, styles, hyperlinks, and formulas together", async () => {
    const cells = new Map<string, Partial<Cell>>();
    cells.set("0,0", {
      value: "Bold Link",
      style: { font: { bold: true } },
      hyperlink: { target: "https://example.com" },
    });
    cells.set("2,0", {
      value: 100,
      formula: "SUM(A1:A2)",
    });

    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["Bold Link", "B1"], [50, 50], [100]],
      cells,
      merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }],
    };

    const data = await writeOds({ sheets: [sheet] });
    const wb = await readOds(data, { readStyles: true });

    // Merges
    expect(wb.sheets[0].merges).toHaveLength(1);

    // Hyperlink on merged cell
    const cell00 = wb.sheets[0].cells!.get("0,0");
    expect(cell00!.hyperlink!.target).toBe("https://example.com");
    expect(cell00!.style?.font?.bold).toBe(true);

    // Formula
    const cell20 = wb.sheets[0].cells!.get("2,0");
    expect(cell20!.formula).toBe("SUM(A1:A2)");
  });

  it("multiple sheets with different features", async () => {
    const cells1 = new Map<string, Partial<Cell>>();
    cells1.set("0,0", {
      value: "Link",
      hyperlink: { target: "https://a.com" },
    });

    const cells2 = new Map<string, Partial<Cell>>();
    cells2.set("0,0", {
      value: 10,
      formula: "5+5",
    });

    const data = await writeOds({
      sheets: [
        {
          name: "Links",
          rows: [["Link"]],
          cells: cells1,
        },
        {
          name: "Formulas",
          rows: [[10]],
          cells: cells2,
          merges: [{ startRow: 0, startCol: 0, endRow: 1, endCol: 1 }],
        },
      ],
    });

    const wb = await readOds(data);

    // Sheet 1: hyperlink
    expect(wb.sheets[0].cells!.get("0,0")!.hyperlink!.target).toBe("https://a.com");

    // Sheet 2: formula + merge
    expect(wb.sheets[1].cells!.get("0,0")!.formula).toBe("5+5");
    expect(wb.sheets[1].merges).toHaveLength(1);
  });
});
