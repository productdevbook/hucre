/**
 * Tests for spec fixes:
 * - Issue #98: maxRows option for XLSX read
 * - Issue #108: calcChain.xml not handled during roundtrip
 * - Issue #123: Row height and custom height not written
 * - Issue #120: ODS reader cannot parse text:span, text:s, text:line-break, text:tab
 */
import { describe, it, expect } from "vitest";
import { ZipReader } from "../src/zip/reader";
import { ZipWriter } from "../src/zip/writer";
import { parseXml } from "../src/xml/parser";
import { writeXlsx } from "../src/xlsx/writer";
import { readXlsx } from "../src/xlsx/reader";
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip";
import { readOds } from "../src/ods/reader";
import type { WriteSheet, CellValue, RowDef } from "../src/_types";

// ── Helpers ──────────────────────────────────────────────────────────

const decoder = new TextDecoder("utf-8");
const encoder = new TextEncoder();

async function extractXml(data: Uint8Array, path: string): Promise<string> {
  const zip = new ZipReader(data);
  const raw = await zip.extract(path);
  return decoder.decode(raw);
}

function findChild(el: { children: Array<unknown> }, localName: string): any {
  return el.children.find((c: any) => typeof c !== "string" && (c.local || c.tag) === localName);
}

function findChildren(el: { children: Array<unknown> }, localName: string): any[] {
  return el.children.filter((c: any) => typeof c !== "string" && (c.local || c.tag) === localName);
}

/**
 * Inject extra ZIP entries into an existing XLSX archive.
 */
async function injectEntries(
  original: Uint8Array,
  extras: Array<{ path: string; data: Uint8Array }>,
): Promise<Uint8Array> {
  const zip = new ZipReader(original);
  const writer = new ZipWriter();

  for (const path of zip.entries()) {
    const data = await zip.extract(path);
    writer.add(path, data, { compress: false });
  }

  for (const entry of extras) {
    writer.add(entry.path, entry.data, { compress: false });
  }

  return writer.build();
}

/**
 * Create a minimal valid ODS file with custom content.xml.
 */
async function createOdsWithContent(contentXml: string): Promise<Uint8Array> {
  const zip = new ZipWriter();

  zip.add("mimetype", encoder.encode("application/vnd.oasis.opendocument.spreadsheet"), {
    compress: false,
  });

  zip.add("content.xml", encoder.encode(contentXml));

  zip.add(
    "META-INF/manifest.xml",
    encoder.encode(
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">' +
        '<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>' +
        '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>' +
        "</manifest:manifest>",
    ),
  );

  return zip.build();
}

// ── Issue #98: maxRows option for XLSX ──────────────────────────────

describe("Issue #98: maxRows option", () => {
  it("limits the number of rows read from XLSX", async () => {
    // Write 100 rows
    const rows: CellValue[][] = [];
    for (let i = 0; i < 100; i++) {
      rows.push([`Row ${i + 1}`, i + 1]);
    }

    const xlsx = await writeXlsx({
      sheets: [{ name: "Sheet1", rows }],
    });

    // Read with maxRows: 10
    const workbook = await readXlsx(xlsx, { maxRows: 10 });
    expect(workbook.sheets).toHaveLength(1);
    expect(workbook.sheets[0].rows.length).toBe(10);
    expect(workbook.sheets[0].rows[0][0]).toBe("Row 1");
    expect(workbook.sheets[0].rows[9][0]).toBe("Row 10");
  });

  it("reads all rows when maxRows is not set", async () => {
    const rows: CellValue[][] = [];
    for (let i = 0; i < 50; i++) {
      rows.push([`Row ${i + 1}`]);
    }

    const xlsx = await writeXlsx({
      sheets: [{ name: "Sheet1", rows }],
    });

    const workbook = await readXlsx(xlsx);
    expect(workbook.sheets[0].rows.length).toBe(50);
  });

  it("reads all rows when maxRows exceeds actual row count", async () => {
    const rows: CellValue[][] = [];
    for (let i = 0; i < 5; i++) {
      rows.push([`Row ${i + 1}`]);
    }

    const xlsx = await writeXlsx({
      sheets: [{ name: "Sheet1", rows }],
    });

    const workbook = await readXlsx(xlsx, { maxRows: 100 });
    expect(workbook.sheets[0].rows.length).toBe(5);
  });
});

// ── Issue #108: calcChain.xml removal ───────────────────────────────

describe("Issue #108: calcChain.xml removal during roundtrip", () => {
  it("removes xl/calcChain.xml from the output ZIP", async () => {
    // Create a basic XLSX
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            ["A", "B"],
            [1, 2],
          ],
        },
      ],
    });

    // Inject a fake calcChain.xml
    const calcChainXml =
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<c r="A1" i="1"/>' +
      "</calcChain>";

    const xlsxWithCalcChain = await injectEntries(xlsx, [
      { path: "xl/calcChain.xml", data: encoder.encode(calcChainXml) },
    ]);

    // Verify calcChain.xml exists in injected file
    const zipBefore = new ZipReader(xlsxWithCalcChain);
    expect(zipBefore.has("xl/calcChain.xml")).toBe(true);

    // Open and save via roundtrip
    const workbook = await openXlsx(xlsxWithCalcChain);
    const saved = await saveXlsx(workbook);

    // Verify calcChain.xml is removed
    const zipAfter = new ZipReader(saved);
    expect(zipAfter.has("xl/calcChain.xml")).toBe(false);
  });

  it("roundtrip works even without calcChain.xml", async () => {
    const xlsx = await writeXlsx({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    });

    const workbook = await openXlsx(xlsx);
    const saved = await saveXlsx(workbook);

    // Should still produce valid XLSX
    const result = await readXlsx(saved);
    expect(result.sheets[0].rows[0][0]).toBe("test");
  });
});

// ── Issue #123: Row height and hidden not written ───────────────────

describe("Issue #123: Row height and custom height", () => {
  it("writes ht and customHeight attributes on rows with height", async () => {
    const rowDefs = new Map<number, RowDef>();
    rowDefs.set(1, { height: 30 });
    rowDefs.set(3, { height: 45.5 });

    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Header"], ["Row 2"], ["Row 3"], ["Row 4"]],
          rowDefs,
        },
      ],
    });

    const sheetXml = await extractXml(xlsx, "xl/worksheets/sheet1.xml");
    const doc = parseXml(sheetXml);
    const sheetData = findChild(doc, "sheetData");
    const xmlRows = findChildren(sheetData, "row");

    // Row 2 (r="2") should have ht="30" customHeight="1"
    const row2 = xmlRows.find((r: any) => r.attrs["r"] === "2");
    expect(row2).toBeDefined();
    expect(row2.attrs["ht"]).toBe("30");
    expect(row2.attrs["customHeight"]).toBe("1");

    // Row 4 (r="4") should have ht="45.5" customHeight="1"
    const row4 = xmlRows.find((r: any) => r.attrs["r"] === "4");
    expect(row4).toBeDefined();
    expect(row4.attrs["ht"]).toBe("45.5");
    expect(row4.attrs["customHeight"]).toBe("1");

    // Row 1 (r="1") should NOT have ht attribute
    const row1 = xmlRows.find((r: any) => r.attrs["r"] === "1");
    expect(row1).toBeDefined();
    expect(row1.attrs["ht"]).toBeUndefined();
  });

  it("writes hidden attribute on hidden rows", async () => {
    const rowDefs = new Map<number, RowDef>();
    rowDefs.set(1, { hidden: true });

    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Visible"], ["Hidden"], ["Visible"]],
          rowDefs,
        },
      ],
    });

    const sheetXml = await extractXml(xlsx, "xl/worksheets/sheet1.xml");
    const doc = parseXml(sheetXml);
    const sheetData = findChild(doc, "sheetData");
    const xmlRows = findChildren(sheetData, "row");

    // Row 2 (r="2") should have hidden="1"
    const row2 = xmlRows.find((r: any) => r.attrs["r"] === "2");
    expect(row2).toBeDefined();
    expect(row2.attrs["hidden"]).toBe("1");

    // Row 1 (r="1") should NOT have hidden attribute
    const row1 = xmlRows.find((r: any) => r.attrs["r"] === "1");
    expect(row1).toBeDefined();
    expect(row1.attrs["hidden"]).toBeUndefined();
  });

  it("round-trips row height: write → read → verify height preserved", async () => {
    const rowDefs = new Map<number, RowDef>();
    rowDefs.set(0, { height: 25 });
    rowDefs.set(2, { height: 40 });

    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["A"], ["B"], ["C"]],
          rowDefs,
        },
      ],
    });

    const workbook = await readXlsx(xlsx);
    const sheet = workbook.sheets[0];

    expect(sheet.rowDefs).toBeDefined();
    expect(sheet.rowDefs!.get(0)?.height).toBe(25);
    expect(sheet.rowDefs!.get(2)?.height).toBe(40);
    // Row 1 should not have a rowDef (no custom height)
    expect(sheet.rowDefs!.has(1)).toBe(false);
  });

  it("round-trips hidden rows: write → read → verify hidden preserved", async () => {
    const rowDefs = new Map<number, RowDef>();
    rowDefs.set(1, { hidden: true });

    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["A"], ["B"], ["C"]],
          rowDefs,
        },
      ],
    });

    const workbook = await readXlsx(xlsx);
    const sheet = workbook.sheets[0];

    expect(sheet.rowDefs).toBeDefined();
    expect(sheet.rowDefs!.get(1)?.hidden).toBe(true);
  });

  it("row height roundtrip through openXlsx/saveXlsx preserves heights", async () => {
    const rowDefs = new Map<number, RowDef>();
    rowDefs.set(0, { height: 20 });
    rowDefs.set(2, { height: 50 });

    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["A"], ["B"], ["C"]],
          rowDefs,
        },
      ],
    });

    // Roundtrip
    const workbook = await openXlsx(xlsx);
    const saved = await saveXlsx(workbook);

    // Read back
    const result = await readXlsx(saved);
    const sheet = result.sheets[0];
    expect(sheet.rowDefs).toBeDefined();
    expect(sheet.rowDefs!.get(0)?.height).toBe(20);
    expect(sheet.rowDefs!.get(2)?.height).toBe(50);
  });
});

// ── Issue #120: ODS text:span, text:s, text:line-break, text:tab ────

describe("Issue #120: ODS inline text elements", () => {
  const NS_OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
  const NS_TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
  const NS_TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

  function odsContentXml(tableRows: string): string {
    return (
      '<?xml version="1.0" encoding="UTF-8"?>' +
      `<office:document-content xmlns:office="${NS_OFFICE}" xmlns:table="${NS_TABLE}" xmlns:text="${NS_TEXT}" office:version="1.2">` +
      "<office:body>" +
      "<office:spreadsheet>" +
      '<table:table table:name="Sheet1">' +
      tableRows +
      "</table:table>" +
      "</office:spreadsheet>" +
      "</office:body>" +
      "</office:document-content>"
    );
  }

  it("parses text:span elements", async () => {
    const contentXml = odsContentXml(
      "<table:table-row>" +
        '<table:table-cell office:value-type="string">' +
        "<text:p>Hello <text:span>World</text:span></text:p>" +
        "</table:table-cell>" +
        "</table:table-row>",
    );

    const ods = await createOdsWithContent(contentXml);
    const workbook = await readOds(ods);

    expect(workbook.sheets[0].rows[0][0]).toBe("Hello World");
  });

  it("parses text:s space elements (default count=1)", async () => {
    const contentXml = odsContentXml(
      "<table:table-row>" +
        '<table:table-cell office:value-type="string">' +
        "<text:p>Hello<text:s/>World</text:p>" +
        "</table:table-cell>" +
        "</table:table-row>",
    );

    const ods = await createOdsWithContent(contentXml);
    const workbook = await readOds(ods);

    expect(workbook.sheets[0].rows[0][0]).toBe("Hello World");
  });

  it("parses text:s space elements with text:c count attribute", async () => {
    const contentXml = odsContentXml(
      "<table:table-row>" +
        '<table:table-cell office:value-type="string">' +
        '<text:p>A<text:s text:c="3"/>B</text:p>' +
        "</table:table-cell>" +
        "</table:table-row>",
    );

    const ods = await createOdsWithContent(contentXml);
    const workbook = await readOds(ods);

    expect(workbook.sheets[0].rows[0][0]).toBe("A   B");
  });

  it("parses text:line-break elements", async () => {
    const contentXml = odsContentXml(
      "<table:table-row>" +
        '<table:table-cell office:value-type="string">' +
        "<text:p>Line1<text:line-break/>Line2</text:p>" +
        "</table:table-cell>" +
        "</table:table-row>",
    );

    const ods = await createOdsWithContent(contentXml);
    const workbook = await readOds(ods);

    expect(workbook.sheets[0].rows[0][0]).toBe("Line1\nLine2");
  });

  it("parses text:tab elements", async () => {
    const contentXml = odsContentXml(
      "<table:table-row>" +
        '<table:table-cell office:value-type="string">' +
        "<text:p>Col1<text:tab/>Col2</text:p>" +
        "</table:table-cell>" +
        "</table:table-row>",
    );

    const ods = await createOdsWithContent(contentXml);
    const workbook = await readOds(ods);

    expect(workbook.sheets[0].rows[0][0]).toBe("Col1\tCol2");
  });

  it("parses mixed inline elements", async () => {
    const contentXml = odsContentXml(
      "<table:table-row>" +
        '<table:table-cell office:value-type="string">' +
        "<text:p>" +
        "<text:span>Bold</text:span>" +
        "<text:s/>" +
        "text" +
        "<text:line-break/>" +
        "next" +
        "<text:tab/>" +
        "tab" +
        '<text:s text:c="2"/>' +
        "end" +
        "</text:p>" +
        "</table:table-cell>" +
        "</table:table-row>",
    );

    const ods = await createOdsWithContent(contentXml);
    const workbook = await readOds(ods);

    expect(workbook.sheets[0].rows[0][0]).toBe("Bold text\nnext\ttab  end");
  });

  it("parses nested text:span elements", async () => {
    const contentXml = odsContentXml(
      "<table:table-row>" +
        '<table:table-cell office:value-type="string">' +
        "<text:p><text:span>outer <text:span>inner</text:span> end</text:span></text:p>" +
        "</table:table-cell>" +
        "</table:table-row>",
    );

    const ods = await createOdsWithContent(contentXml);
    const workbook = await readOds(ods);

    expect(workbook.sheets[0].rows[0][0]).toBe("outer inner end");
  });
});
