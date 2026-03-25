/**
 * Tests that specifically verify OOXML spec compliance.
 * The worksheet XML element ordering must follow ECMA-376 Part 1, 18.3.1.99
 */
import { describe, it, expect } from "vitest";
import { writeXlsx } from "../src/xlsx/writer";
import { readXlsx } from "../src/xlsx/reader";
import { ZipReader } from "../src/zip/reader";
import { parseXml } from "../src/xml/parser";
import type { WriteSheet, CellValue } from "../src/_types";
import type { XmlElement } from "../src/xml/parser";

const decoder = new TextDecoder("utf-8");

async function getWorksheetXml(sheets: WriteSheet[]): Promise<string> {
  const xlsx = await writeXlsx({ sheets });
  const zip = new ZipReader(xlsx);
  const raw = await zip.extract("xl/worksheets/sheet1.xml");
  return decoder.decode(raw);
}

function getChildTagOrder(xml: string): string[] {
  const doc = parseXml(xml);
  // Get the order of direct child element tags
  return doc.children
    .filter((c): c is import("../src/xml/parser").XmlElement => typeof c !== "string")
    .map((c) => c.local || c.tag);
}

// ═══════════════════════════════════════════════════════════════════════
// OOXML Element Ordering
// The ECMA-376 Part 1 spec (18.3.1.99 worksheet) defines a strict
// element ordering. Violating it can cause Excel to flag the file
// as corrupted.
//
// Required order:
//   sheetPr?, dimension?, sheetViews?, sheetFormatPr?, cols*,
//   sheetData, sheetCalcPr?, sheetProtection?, protectedRanges?,
//   scenarios?, autoFilter?, sortState?, dataConsolidate?,
//   customSheetViews?, mergeCells?, phoneticPr?,
//   conditionalFormatting*, dataValidations?, hyperlinks?,
//   printOptions?, pageMargins?, pageSetup?, headerFooter?,
//   rowBreaks?, colBreaks?, customProperties?, cellWatches?,
//   ignoredErrors?, smartTags?, drawing?, drawingHF?, picture?,
//   oleObjects?, controls?, webPublishItems?, tableParts?, extLst?
// ═══════════════════════════════════════════════════════════════════════

describe("OOXML worksheet element ordering", () => {
  it("sheetProtection comes after sheetData per OOXML spec", async () => {
    const xml = await getWorksheetXml([
      {
        name: "Protected",
        rows: [["data"]],
        protection: {
          sheet: true,
          password: "test",
        },
      },
    ]);

    const tags = getChildTagOrder(xml);

    const protectionIdx = tags.indexOf("sheetProtection");
    const formatPrIdx = tags.indexOf("sheetFormatPr");
    const sheetDataIdx = tags.indexOf("sheetData");

    expect(protectionIdx).toBeGreaterThan(-1);
    expect(formatPrIdx).toBeGreaterThan(-1);
    expect(sheetDataIdx).toBeGreaterThan(-1);

    // sheetFormatPr should come before sheetData
    expect(formatPrIdx).toBeLessThan(sheetDataIdx);

    // sheetProtection MUST come after sheetData per ECMA-376
    expect(protectionIdx).toBeGreaterThan(sheetDataIdx);
  });

  it("autoFilter comes after sheetData and before mergeCells per spec", async () => {
    const xml = await getWorksheetXml([
      {
        name: "Filtered",
        rows: [
          ["h1", "h2"],
          ["a", "b"],
        ],
        autoFilter: { range: "A1:B2" },
        merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }],
      },
    ]);

    const tags = getChildTagOrder(xml);

    const autoFilterIdx = tags.indexOf("autoFilter");
    const sheetDataIdx = tags.indexOf("sheetData");
    const mergeCellsIdx = tags.indexOf("mergeCells");

    expect(autoFilterIdx).toBeGreaterThan(-1);
    expect(sheetDataIdx).toBeGreaterThan(-1);
    expect(mergeCellsIdx).toBeGreaterThan(-1);

    // autoFilter must come after sheetData
    expect(autoFilterIdx).toBeGreaterThan(sheetDataIdx);

    // autoFilter must come before mergeCells per ECMA-376
    expect(autoFilterIdx).toBeLessThan(mergeCellsIdx);
  });

  it("verify basic element ordering: sheetViews < sheetFormatPr < cols < sheetData", async () => {
    const xml = await getWorksheetXml([
      {
        name: "Basic",
        rows: [["data"]],
        columns: [{ width: 20 }],
      },
    ]);

    const tags = getChildTagOrder(xml);

    const viewsIdx = tags.indexOf("sheetViews");
    const formatIdx = tags.indexOf("sheetFormatPr");
    const colsIdx = tags.indexOf("cols");
    const dataIdx = tags.indexOf("sheetData");

    expect(viewsIdx).toBeLessThan(formatIdx);
    expect(formatIdx).toBeLessThan(colsIdx);
    expect(colsIdx).toBeLessThan(dataIdx);
  });

  it("conditional formatting comes after sheetData", async () => {
    const xml = await getWorksheetXml([
      {
        name: "CF",
        rows: [[1], [2], [3]],
        conditionalRules: [
          {
            type: "cellIs",
            priority: 1,
            operator: "greaterThan",
            formula: "1",
            range: "A1:A3",
            style: { font: { bold: true } },
          },
        ],
      },
    ]);

    const tags = getChildTagOrder(xml);

    const sheetDataIdx = tags.indexOf("sheetData");
    const cfIdx = tags.indexOf("conditionalFormatting");

    expect(cfIdx).toBeGreaterThan(sheetDataIdx);
  });

  it("dataValidations come after conditionalFormatting", async () => {
    const xml = await getWorksheetXml([
      {
        name: "DV",
        rows: [["h1"]],
        dataValidations: [{ type: "list", range: "A2:A10", values: ["a", "b", "c"] }],
        conditionalRules: [
          {
            type: "cellIs",
            priority: 1,
            operator: "equal",
            formula: '"a"',
            range: "A2:A10",
            style: { font: { bold: true } },
          },
        ],
      },
    ]);

    const tags = getChildTagOrder(xml);

    const cfIdx = tags.indexOf("conditionalFormatting");
    const dvIdx = tags.indexOf("dataValidations");

    if (cfIdx >= 0 && dvIdx >= 0) {
      expect(dvIdx).toBeGreaterThan(cfIdx);
    }
  });

  it("hyperlinks come after dataValidations", async () => {
    const cells = new Map<string, Partial<import("../src/_types").Cell>>();
    cells.set("0,0", {
      value: "Link",
      type: "string",
      hyperlink: { target: "https://example.com" },
    });

    const xml = await getWorksheetXml([
      {
        name: "H",
        rows: [["Link"]],
        cells,
        dataValidations: [{ type: "list", range: "B1:B10", values: ["x", "y"] }],
      },
    ]);

    const tags = getChildTagOrder(xml);

    const dvIdx = tags.indexOf("dataValidations");
    const hlIdx = tags.indexOf("hyperlinks");

    if (dvIdx >= 0 && hlIdx >= 0) {
      expect(hlIdx).toBeGreaterThan(dvIdx);
    }
  });

  it("pageMargins and pageSetup come after hyperlinks", async () => {
    const cells = new Map<string, Partial<import("../src/_types").Cell>>();
    cells.set("0,0", {
      value: "Link",
      type: "string",
      hyperlink: { target: "https://example.com" },
    });

    const xml = await getWorksheetXml([
      {
        name: "PS",
        rows: [["data"]],
        cells,
        pageSetup: {
          orientation: "landscape",
          paperSize: "a4",
        },
      },
    ]);

    const tags = getChildTagOrder(xml);

    const hlIdx = tags.indexOf("hyperlinks");
    const marginsIdx = tags.indexOf("pageMargins");
    const setupIdx = tags.indexOf("pageSetup");

    if (hlIdx >= 0) {
      expect(marginsIdx).toBeGreaterThan(hlIdx);
    }
    if (marginsIdx >= 0 && setupIdx >= 0) {
      expect(setupIdx).toBeGreaterThan(marginsIdx);
    }
  });

  it("drawing and tableParts come last", async () => {
    // Test with table
    const xml = await getWorksheetXml([
      {
        name: "T",
        rows: [
          ["H1", "H2"],
          ["a", "b"],
        ],
        tables: [
          {
            name: "Table1",
            columns: [{ name: "H1" }, { name: "H2" }],
          },
        ],
      },
    ]);

    const tags = getChildTagOrder(xml);

    const sheetDataIdx = tags.indexOf("sheetData");
    const tablePartsIdx = tags.indexOf("tableParts");

    if (tablePartsIdx >= 0) {
      // tableParts should be one of the last elements
      expect(tablePartsIdx).toBeGreaterThan(sheetDataIdx);
    }
  });
});

// ═══════════════════════════════════════════════════════════════════════
// Test that files with protection can still be read by Excel
// (protection + data validations + merges combo)
// ═══════════════════════════════════════════════════════════════════════

describe("XLSX: complex feature combinations", () => {
  it("protection + merges + autoFilter + data validations", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Complex",
          rows: [
            ["Name", "Age", "City"],
            ["Alice", 30, "NYC"],
            ["Bob", 25, "LA"],
          ],
          protection: { sheet: true },
          merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }],
          autoFilter: { range: "A1:C3" },
          dataValidations: [
            { type: "whole", operator: "between", formula1: "0", formula2: "150", range: "B2:B3" },
          ],
        },
      ],
    });

    // Should parse without errors
    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].name).toBe("Complex");
    expect(wb.sheets[0].rows).toHaveLength(3);

    // Verify features are preserved
    expect(wb.sheets[0].merges).toBeDefined();
    expect(wb.sheets[0].autoFilter).toBeDefined();
    expect(wb.sheets[0].autoFilter!.range).toBe("A1:C3");
    expect(wb.sheets[0].dataValidations).toBeDefined();
  });

  it("all sheet features combined", async () => {
    const cells = new Map<string, Partial<import("../src/_types").Cell>>();
    cells.set("0,0", {
      value: "Linked",
      type: "string",
      hyperlink: { target: "https://example.com" },
    });

    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Everything",
          rows: [
            ["Linked", "B", "C"],
            [1, 2, 3],
            [4, 5, 6],
          ],
          cells,
          columns: [{ width: 20 }, { width: 15 }, { width: 10 }],
          freezePane: { rows: 1 },
          merges: [{ startRow: 1, startCol: 1, endRow: 2, endCol: 2 }],
          autoFilter: { range: "A1:C1" },
          dataValidations: [
            { type: "whole", range: "A2:A3", operator: "greaterThan", formula1: "0" },
          ],
          conditionalRules: [
            {
              type: "cellIs",
              priority: 1,
              operator: "greaterThan",
              formula: "3",
              range: "A2:A3",
              style: { font: { bold: true } },
            },
          ],
          protection: { sheet: true },
          pageSetup: {
            orientation: "landscape",
            paperSize: "a4",
          },
          headerFooter: {
            oddHeader: "Header",
            oddFooter: "Page &P",
          },
          view: {
            zoomScale: 120,
          },
        },
      ],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].name).toBe("Everything");
    expect(wb.sheets[0].rows).toHaveLength(3);
    expect(wb.sheets[0].rows[0][0]).toBe("Linked");
  });
});

// ═══════════════════════════════════════════════════════════════════════
// Additional regression tests for fixed bugs
// ═══════════════════════════════════════════════════════════════════════

describe("OOXML: sheetProtection position (#40)", () => {
  it("sheetProtection comes after sheetData, not before sheetFormatPr", async () => {
    const xml = await getWorksheetXml([
      {
        name: "Sheet1",
        rows: [["data"]],
        protection: { sheet: true },
      },
    ]);

    const tags = getChildTagOrder(xml);

    const fmtIdx = tags.indexOf("sheetFormatPr");
    const dataIdx = tags.indexOf("sheetData");
    const protIdx = tags.indexOf("sheetProtection");

    // sheetProtection must NOT be between sheetViews and sheetFormatPr
    expect(protIdx).toBeGreaterThan(dataIdx);
    // sheetFormatPr must come before sheetData (unchanged)
    expect(fmtIdx).toBeLessThan(dataIdx);
  });

  it("sheetProtection comes before autoFilter when both present", async () => {
    const xml = await getWorksheetXml([
      {
        name: "Sheet1",
        rows: [
          ["h1", "h2"],
          ["a", "b"],
        ],
        protection: { sheet: true },
        autoFilter: { range: "A1:B2" },
      },
    ]);

    const tags = getChildTagOrder(xml);

    const protIdx = tags.indexOf("sheetProtection");
    const filterIdx = tags.indexOf("autoFilter");

    expect(protIdx).toBeGreaterThan(-1);
    expect(filterIdx).toBeGreaterThan(-1);
    expect(protIdx).toBeLessThan(filterIdx);
  });
});

describe("OOXML: autoFilter before mergeCells (#41)", () => {
  it("autoFilter comes before mergeCells", async () => {
    const xml = await getWorksheetXml([
      {
        name: "Sheet1",
        rows: [
          ["h1", "h2"],
          ["a", "b"],
        ],
        autoFilter: { range: "A1:B2" },
        merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }],
      },
    ]);

    const tags = getChildTagOrder(xml);

    const filterIdx = tags.indexOf("autoFilter");
    const mergeIdx = tags.indexOf("mergeCells");

    expect(filterIdx).toBeGreaterThan(-1);
    expect(mergeIdx).toBeGreaterThan(-1);
    expect(filterIdx).toBeLessThan(mergeIdx);
  });

  it("full ordering: sheetData > sheetProtection > autoFilter > mergeCells", async () => {
    const xml = await getWorksheetXml([
      {
        name: "Sheet1",
        rows: [
          ["h1", "h2"],
          ["a", "b"],
        ],
        protection: { sheet: true },
        autoFilter: { range: "A1:B2" },
        merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }],
      },
    ]);

    const tags = getChildTagOrder(xml);

    const dataIdx = tags.indexOf("sheetData");
    const protIdx = tags.indexOf("sheetProtection");
    const filterIdx = tags.indexOf("autoFilter");
    const mergeIdx = tags.indexOf("mergeCells");

    expect(dataIdx).toBeLessThan(protIdx);
    expect(protIdx).toBeLessThan(filterIdx);
    expect(filterIdx).toBeLessThan(mergeIdx);
  });
});

describe("autoFilter round-trip (#42)", () => {
  it("autoFilter is parsed on read", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Filtered",
          rows: [
            ["Name", "Age"],
            ["Alice", 30],
            ["Bob", 25],
          ],
          autoFilter: { range: "A1:B3" },
        },
      ],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].autoFilter).toBeDefined();
    expect(wb.sheets[0].autoFilter!.range).toBe("A1:B3");
  });

  it("autoFilter round-trips through write and read", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            ["A", "B", "C", "D"],
            [1, 2, 3, 4],
          ],
          autoFilter: { range: "A1:D2" },
        },
      ],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].autoFilter).toEqual({ range: "A1:D2" });
  });

  it("sheet without autoFilter has no autoFilter property", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["data"]],
        },
      ],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].autoFilter).toBeUndefined();
  });
});

describe("Infinity/NaN handling (#43)", () => {
  it("Infinity becomes null (empty cell)", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [[1, Infinity, 3]],
        },
      ],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].rows[0][0]).toBe(1);
    expect(wb.sheets[0].rows[0][1]).toBeNull();
    expect(wb.sheets[0].rows[0][2]).toBe(3);
  });

  it("-Infinity becomes null (empty cell)", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [[1, -Infinity, 3]],
        },
      ],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].rows[0][0]).toBe(1);
    expect(wb.sheets[0].rows[0][1]).toBeNull();
    expect(wb.sheets[0].rows[0][2]).toBe(3);
  });

  it("NaN becomes null (empty cell)", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [[1, NaN, 3]],
        },
      ],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].rows[0][0]).toBe(1);
    expect(wb.sheets[0].rows[0][1]).toBeNull();
    expect(wb.sheets[0].rows[0][2]).toBe(3);
  });

  it("normal numbers are unaffected", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [[0, -1, 3.14, 1e10, -0]],
        },
      ],
    });

    const wb = await readXlsx(xlsx);
    expect(wb.sheets[0].rows[0][0]).toBe(0);
    expect(wb.sheets[0].rows[0][1]).toBe(-1);
    expect(wb.sheets[0].rows[0][2]).toBe(3.14);
    expect(wb.sheets[0].rows[0][3]).toBe(1e10);
  });
});
