import { describe, it, expect } from "vitest";
import { ZipReader } from "../src/zip/reader";
import { parseXml } from "../src/xml/parser";
import { writeOds } from "../src/ods/writer";
import { readOds } from "../src/ods/reader";
import type { CellValue } from "../src/_types";

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

function getElementText(el: { children: Array<unknown> }): string {
  return el.children.filter((c: unknown) => typeof c === "string").join("");
}

/** Navigate content.xml to get the first table element */
async function getFirstTable(data: Uint8Array) {
  const contentDoc = await parseXmlFromZip(data, "content.xml");
  const body = findChild(contentDoc, "body");
  const spreadsheet = findChild(body, "spreadsheet");
  return findChild(spreadsheet, "table");
}

// ── ODS Writer Tests ────────────────────────────────────────────────

describe("ODS Writer", () => {
  it("produces a valid ZIP archive", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["Hello"]] }],
    });

    const zip = new ZipReader(data);
    expect(zip.has("mimetype")).toBe(true);
    expect(zip.has("content.xml")).toBe(true);
    expect(zip.has("meta.xml")).toBe(true);
    expect(zip.has("styles.xml")).toBe(true);
    expect(zip.has("META-INF/manifest.xml")).toBe(true);
  });

  it("writes mimetype as the first entry and uncompressed", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    });

    const zip = new ZipReader(data);
    const entries = zip.entries();

    // mimetype should be the first entry
    expect(entries[0]).toBe("mimetype");

    // mimetype content should be the correct MIME type
    const mimeData = await zip.extract("mimetype");
    expect(decoder.decode(mimeData)).toBe("application/vnd.oasis.opendocument.spreadsheet");

    // Verify it's stored uncompressed by checking compression method in local file header
    // The local file header starts at offset 0 for the first entry
    // Compression method is at offset 8 in the local file header
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    const compressionMethod = view.getUint16(8, true);
    expect(compressionMethod).toBe(0); // STORE = 0
  });

  it("writes mimetype entry with no extra field in local file header", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    });

    // The local file header for mimetype starts at offset 0
    // Extra field length is at offset 28 in the local file header
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    const extraFieldLength = view.getUint16(28, true);
    expect(extraFieldLength).toBe(0);
  });

  it("writes manifest.xml with required entries", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    });

    const manifest = await parseXmlFromZip(data, "META-INF/manifest.xml");
    const entries = findChildren(manifest, "file-entry");

    // Should have entries for /, content.xml, meta.xml, styles.xml
    const paths = entries.map((e: any) => e.attrs["manifest:full-path"]);
    expect(paths).toContain("/");
    expect(paths).toContain("content.xml");
    expect(paths).toContain("meta.xml");
    expect(paths).toContain("styles.xml");
  });

  it("writes manifest.xml with correct media types", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    });

    const manifest = await parseXmlFromZip(data, "META-INF/manifest.xml");
    const entries = findChildren(manifest, "file-entry");

    const entryMap = new Map(
      entries.map((e: any) => [e.attrs["manifest:full-path"], e.attrs["manifest:media-type"]]),
    );
    expect(entryMap.get("/")).toBe("application/vnd.oasis.opendocument.spreadsheet");
    expect(entryMap.get("content.xml")).toBe("text/xml");
    expect(entryMap.get("meta.xml")).toBe("text/xml");
    expect(entryMap.get("styles.xml")).toBe("text/xml");
  });

  it("writes manifest.xml with version 1.2 on root entry", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    });

    const manifest = await parseXmlFromZip(data, "META-INF/manifest.xml");
    const entries = findChildren(manifest, "file-entry");
    const rootEntry = entries.find((e: any) => e.attrs["manifest:full-path"] === "/");
    expect(rootEntry.attrs["manifest:version"]).toBe("1.2");
  });

  it("writes string cells correctly", async () => {
    const data = await writeOds({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            ["Hello", "World"],
            ["Foo", "Bar"],
          ],
        },
      ],
    });

    const contentDoc = await parseXmlFromZip(data, "content.xml");
    const body = findChild(contentDoc, "body");
    const spreadsheet = findChild(body, "spreadsheet");
    const table = findChild(spreadsheet, "table");

    expect(table.attrs["table:name"]).toBe("Sheet1");

    const tableRows = findChildren(table, "table-row");
    expect(tableRows.length).toBe(2);

    // First row, first cell
    const firstRowCells = findChildren(tableRows[0], "table-cell");
    expect(firstRowCells[0].attrs["office:value-type"]).toBe("string");
    const textP = findChild(firstRowCells[0], "p");
    expect(getElementText(textP)).toBe("Hello");
  });

  it("writes numeric cells correctly", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [[42, 3.14]] }],
    });

    const contentDoc = await parseXmlFromZip(data, "content.xml");
    const body = findChild(contentDoc, "body");
    const spreadsheet = findChild(body, "spreadsheet");
    const table = findChild(spreadsheet, "table");
    const row = findChildren(table, "table-row")[0];
    const cells = findChildren(row, "table-cell");

    expect(cells[0].attrs["office:value-type"]).toBe("float");
    expect(cells[0].attrs["office:value"]).toBe("42");

    expect(cells[1].attrs["office:value-type"]).toBe("float");
    expect(cells[1].attrs["office:value"]).toBe("3.14");
  });

  it("writes boolean cells with office:boolean-value attribute", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [[true, false]] }],
    });

    const contentDoc = await parseXmlFromZip(data, "content.xml");
    const body = findChild(contentDoc, "body");
    const spreadsheet = findChild(body, "spreadsheet");
    const table = findChild(spreadsheet, "table");
    const row = findChildren(table, "table-row")[0];
    const cells = findChildren(row, "table-cell");

    // Verify office:value-type is "boolean"
    expect(cells[0].attrs["office:value-type"]).toBe("boolean");
    expect(cells[1].attrs["office:value-type"]).toBe("boolean");

    // Verify office:boolean-value (NOT office:value) is used
    expect(cells[0].attrs["office:boolean-value"]).toBe("true");
    expect(cells[1].attrs["office:boolean-value"]).toBe("false");

    // Verify office:value is NOT set (only office:boolean-value)
    expect(cells[0].attrs["office:value"]).toBeUndefined();
    expect(cells[1].attrs["office:value"]).toBeUndefined();

    // Verify text:p contains display text
    expect(getElementText(findChild(cells[0], "p"))).toBe("TRUE");
    expect(getElementText(findChild(cells[1], "p"))).toBe("FALSE");
  });

  it("writes date cells with office:date-value attribute in ISO 8601 format", async () => {
    const date = new Date(2026, 2, 24, 10, 30, 0); // March 24, 2026
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [[date]] }],
    });

    const contentDoc = await parseXmlFromZip(data, "content.xml");
    const body = findChild(contentDoc, "body");
    const spreadsheet = findChild(body, "spreadsheet");
    const table = findChild(spreadsheet, "table");
    const row = findChildren(table, "table-row")[0];
    const cells = findChildren(row, "table-cell");

    expect(cells[0].attrs["office:value-type"]).toBe("date");
    expect(cells[0].attrs["office:date-value"]).toBe("2026-03-24T10:30:00");

    // Verify text:p is present inside date cell
    const textP = findChild(cells[0], "p");
    expect(textP).toBeDefined();
    expect(getElementText(textP)).toBe("2026-03-24T10:30:00");
  });

  it("writes multiple sheets", async () => {
    const data = await writeOds({
      sheets: [
        { name: "First", rows: [["A"]] },
        { name: "Second", rows: [["B"]] },
        { name: "Third", rows: [["C"]] },
      ],
    });

    const contentDoc = await parseXmlFromZip(data, "content.xml");
    const body = findChild(contentDoc, "body");
    const spreadsheet = findChild(body, "spreadsheet");
    const tables = findChildren(spreadsheet, "table");

    expect(tables.length).toBe(3);
    expect(tables[0].attrs["table:name"]).toBe("First");
    expect(tables[1].attrs["table:name"]).toBe("Second");
    expect(tables[2].attrs["table:name"]).toBe("Third");
  });

  it("writes null cells as empty", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["A", null, "C"]] }],
    });

    const contentDoc = await parseXmlFromZip(data, "content.xml");
    const body = findChild(contentDoc, "body");
    const spreadsheet = findChild(body, "spreadsheet");
    const table = findChild(spreadsheet, "table");
    const row = findChildren(table, "table-row")[0];
    const cells = findChildren(row, "table-cell");

    // Middle cell should be empty (self-closing)
    expect(cells[1].attrs["office:value-type"]).toBeUndefined();
  });

  it("writes Application: defter in meta.xml", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    });

    const metaDoc = await parseXmlFromZip(data, "meta.xml");
    const metaEl = findChild(metaDoc, "meta");
    const generator = findChild(metaEl, "generator");
    expect(getElementText(generator)).toBe("defter");
  });

  it("writes document properties in meta.xml", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
      properties: {
        title: "Test Title",
        creator: "Test Author",
        subject: "Test Subject",
      },
    });

    const metaDoc = await parseXmlFromZip(data, "meta.xml");
    const metaEl = findChild(metaDoc, "meta");

    expect(getElementText(findChild(metaEl, "title"))).toBe("Test Title");
    expect(getElementText(findChild(metaEl, "initial-creator"))).toBe("Test Author");
    expect(getElementText(findChild(metaEl, "subject"))).toBe("Test Subject");
  });

  // ── Spec Compliance Tests ────────────────────────────────────────

  it("content.xml has office:scripts element", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    });

    const contentDoc = await parseXmlFromZip(data, "content.xml");
    const scripts = findChild(contentDoc, "scripts");
    expect(scripts).toBeDefined();
  });

  it("content.xml has office:font-face-decls element", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    });

    const contentDoc = await parseXmlFromZip(data, "content.xml");
    const fontFaceDecls = findChild(contentDoc, "font-face-decls");
    expect(fontFaceDecls).toBeDefined();
  });

  it("content.xml has office:automatic-styles element", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    });

    const contentDoc = await parseXmlFromZip(data, "content.xml");
    const autoStyles = findChild(contentDoc, "automatic-styles");
    expect(autoStyles).toBeDefined();
  });

  it("content.xml has office:version attribute", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    });

    const contentDoc = await parseXmlFromZip(data, "content.xml");
    expect(contentDoc.attrs["office:version"]).toBe("1.2");
  });

  it("content.xml has required namespace declarations", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    });

    const contentXml = await extractFile(data, "content.xml");

    // Check all required namespaces are declared
    expect(contentXml).toContain('xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"');
    expect(contentXml).toContain('xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"');
    expect(contentXml).toContain('xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"');
    expect(contentXml).toContain('xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"');
    expect(contentXml).toContain(
      'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
    );
    expect(contentXml).toContain(
      'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"',
    );
    expect(contentXml).toContain(
      'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
    );
  });

  it("content.xml elements appear in correct order per spec", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    });

    const contentDoc = await parseXmlFromZip(data, "content.xml");
    const childElements = contentDoc.children.filter((c: any) => typeof c !== "string");

    // Per ODS spec order: scripts, font-face-decls, automatic-styles, body
    const names = childElements.map((c: any) => c.local || c.tag);
    expect(names).toEqual(["scripts", "font-face-decls", "automatic-styles", "body"]);
  });

  it("table has table:table-column elements", async () => {
    const data = await writeOds({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            ["A", "B", "C"],
            [1, 2, 3],
          ],
        },
      ],
    });

    const table = await getFirstTable(data);
    const columns = findChildren(table, "table-column");

    // Should have table-column element(s) declaring 3 columns
    expect(columns.length).toBeGreaterThanOrEqual(1);

    // The total column count should equal 3
    let totalCols = 0;
    for (const col of columns) {
      const repeat = Number(col.attrs["table:number-columns-repeated"] ?? "1");
      totalCols += repeat;
    }
    expect(totalCols).toBe(3);
  });

  it("table:table-column appears before table:table-row", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["A", "B"]] }],
    });

    const table = await getFirstTable(data);
    const children = table.children.filter((c: any) => typeof c !== "string");
    const names = children.map((c: any) => c.local || c.tag);

    // table-column should come before table-row
    const colIdx = names.indexOf("table-column");
    const rowIdx = names.indexOf("table-row");
    expect(colIdx).toBeLessThan(rowIdx);
  });

  it("uses number-columns-repeated for consecutive empty cells in middle of row", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["A", null, null, null, "B"]] }],
    });

    const table = await getFirstTable(data);
    const row = findChildren(table, "table-row")[0];
    const cells = findChildren(row, "table-cell");

    // Should be: "A" cell, repeated empty cell (3x), "B" cell = 3 elements total
    expect(cells.length).toBe(3);
    expect(cells[1].attrs["table:number-columns-repeated"]).toBe("3");
  });

  it("omits trailing null cells from rows", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["A", "B", null, null, null]] }],
    });

    const table = await getFirstTable(data);
    const row = findChildren(table, "table-row")[0];
    const cells = findChildren(row, "table-cell");

    // Trailing nulls should not be emitted
    expect(cells.length).toBe(2);
    expect(cells[0].attrs["office:value-type"]).toBe("string");
    expect(cells[1].attrs["office:value-type"]).toBe("string");
  });

  it("styles.xml has required child elements", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    });

    const stylesDoc = await parseXmlFromZip(data, "styles.xml");

    // Per ODS spec, styles.xml should have these child elements
    expect(findChild(stylesDoc, "font-face-decls")).toBeDefined();
    expect(findChild(stylesDoc, "styles")).toBeDefined();
    expect(findChild(stylesDoc, "automatic-styles")).toBeDefined();
    expect(findChild(stylesDoc, "master-styles")).toBeDefined();
  });

  it("styles.xml has office:version attribute", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    });

    const stylesDoc = await parseXmlFromZip(data, "styles.xml");
    expect(stylesDoc.attrs["office:version"]).toBe("1.2");
  });

  it("styles.xml has required namespace declarations", async () => {
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    });

    const stylesXml = await extractFile(data, "styles.xml");
    expect(stylesXml).toContain("xmlns:office=");
    expect(stylesXml).toContain("xmlns:style=");
    expect(stylesXml).toContain("xmlns:text=");
    expect(stylesXml).toContain("xmlns:table=");
    expect(stylesXml).toContain("xmlns:fo=");
  });

  it("every non-empty cell has a text:p child element", async () => {
    const date = new Date(2026, 0, 1);
    const data = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["text", 42, true, date]] }],
    });

    const table = await getFirstTable(data);
    const row = findChildren(table, "table-row")[0];
    const cells = findChildren(row, "table-cell");

    for (const cell of cells) {
      const textP = findChild(cell, "p");
      expect(textP).toBeDefined();
      expect(getElementText(textP).length).toBeGreaterThan(0);
    }
  });

  it("empty table has no table-column element", async () => {
    const data = await writeOds({
      sheets: [{ name: "Empty", rows: [] }],
    });

    const contentDoc = await parseXmlFromZip(data, "content.xml");
    const body = findChild(contentDoc, "body");
    const spreadsheet = findChild(body, "spreadsheet");
    const table = findChild(spreadsheet, "table");
    const columns = findChildren(table, "table-column");

    expect(columns.length).toBe(0);
  });
});

// ── ODS Reader Tests ────────────────────────────────────────────────

describe("ODS Reader", () => {
  it("reads back written ODS with string data", async () => {
    const written = await writeOds({
      sheets: [
        {
          name: "Sheet1",
          rows: [
            ["Hello", "World"],
            ["Foo", "Bar"],
          ],
        },
      ],
    });

    const workbook = await readOds(written);
    expect(workbook.sheets.length).toBe(1);
    expect(workbook.sheets[0].name).toBe("Sheet1");
    expect(workbook.sheets[0].rows.length).toBe(2);
    expect(workbook.sheets[0].rows[0]).toEqual(["Hello", "World"]);
    expect(workbook.sheets[0].rows[1]).toEqual(["Foo", "Bar"]);
  });

  it("reads back written ODS with numeric data", async () => {
    const written = await writeOds({
      sheets: [{ name: "Numbers", rows: [[1, 2.5, 100]] }],
    });

    const workbook = await readOds(written);
    expect(workbook.sheets[0].rows[0]).toEqual([1, 2.5, 100]);
  });

  it("reads back written ODS with boolean data", async () => {
    const written = await writeOds({
      sheets: [{ name: "Booleans", rows: [[true, false, true]] }],
    });

    const workbook = await readOds(written);
    expect(workbook.sheets[0].rows[0]).toEqual([true, false, true]);
  });

  it("reads back written ODS with date values", async () => {
    const date1 = new Date("2026-01-15T10:00:00");
    const date2 = new Date("2026-06-30T23:59:59");

    const written = await writeOds({
      sheets: [{ name: "Dates", rows: [[date1, date2]] }],
    });

    const workbook = await readOds(written);
    const row = workbook.sheets[0].rows[0];

    // Dates may not be exactly equal due to time zone handling in ODS,
    // but they should be Date objects with similar values
    expect(row[0]).toBeInstanceOf(Date);
    expect(row[1]).toBeInstanceOf(Date);
  });

  it("reads back written ODS with multiple sheets", async () => {
    const written = await writeOds({
      sheets: [
        { name: "First", rows: [["A1"]] },
        { name: "Second", rows: [["B1", "B2"]] },
        { name: "Third", rows: [["C1"], ["C2"], ["C3"]] },
      ],
    });

    const workbook = await readOds(written);
    expect(workbook.sheets.length).toBe(3);
    expect(workbook.sheets[0].name).toBe("First");
    expect(workbook.sheets[0].rows).toEqual([["A1"]]);
    expect(workbook.sheets[1].name).toBe("Second");
    expect(workbook.sheets[1].rows).toEqual([["B1", "B2"]]);
    expect(workbook.sheets[2].name).toBe("Third");
    expect(workbook.sheets[2].rows).toEqual([["C1"], ["C2"], ["C3"]]);
  });

  it("reads document properties from meta.xml", async () => {
    const written = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
      properties: {
        title: "Test Title",
        creator: "Author",
        subject: "Subject",
        description: "A description",
        keywords: "test, ods",
      },
    });

    const workbook = await readOds(written);
    expect(workbook.properties).toBeDefined();
    expect(workbook.properties!.title).toBe("Test Title");
    expect(workbook.properties!.creator).toBe("Author");
    expect(workbook.properties!.subject).toBe("Subject");
    expect(workbook.properties!.description).toBe("A description");
    expect(workbook.properties!.keywords).toBe("test, ods");
  });

  it("handles mixed data types in a row", async () => {
    const written = await writeOds({
      sheets: [
        {
          name: "Mixed",
          rows: [["text", 42, true, null, "more"]],
        },
      ],
    });

    const workbook = await readOds(written);
    const row = workbook.sheets[0].rows[0];
    expect(row[0]).toBe("text");
    expect(row[1]).toBe(42);
    expect(row[2]).toBe(true);
    // null cells at the end are trimmed, but mid-row null should be preserved
    // However, the writer emits a self-closing cell which has no value-type,
    // reader returns null for such cells
    expect(row[3]).toBe(null);
    expect(row[4]).toBe("more");
  });

  it("handles empty sheets", async () => {
    const written = await writeOds({
      sheets: [{ name: "Empty", rows: [] }],
    });

    const workbook = await readOds(written);
    expect(workbook.sheets.length).toBe(1);
    expect(workbook.sheets[0].name).toBe("Empty");
    expect(workbook.sheets[0].rows).toEqual([]);
  });

  it("rejects non-ZIP data", async () => {
    const badData = new Uint8Array([
      1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22,
    ]);
    await expect(readOds(badData)).rejects.toThrow();
  });

  it("round-trips a complex workbook", async () => {
    const rows: CellValue[][] = [
      ["Name", "Age", "Active", "Score"],
      ["Alice", 30, true, 95.5],
      ["Bob", 25, false, 87.3],
      ["Charlie", 35, true, 92.1],
    ];

    const written = await writeOds({
      sheets: [
        { name: "People", rows },
        { name: "Summary", rows: [["Total", 3]] },
      ],
    });

    const workbook = await readOds(written);

    expect(workbook.sheets.length).toBe(2);

    // First sheet
    expect(workbook.sheets[0].name).toBe("People");
    expect(workbook.sheets[0].rows[0]).toEqual(["Name", "Age", "Active", "Score"]);
    expect(workbook.sheets[0].rows[1]).toEqual(["Alice", 30, true, 95.5]);
    expect(workbook.sheets[0].rows[2]).toEqual(["Bob", 25, false, 87.3]);
    expect(workbook.sheets[0].rows[3]).toEqual(["Charlie", 35, true, 92.1]);

    // Second sheet
    expect(workbook.sheets[1].name).toBe("Summary");
    expect(workbook.sheets[1].rows[0]).toEqual(["Total", 3]);
  });

  it("accepts ArrayBuffer input", async () => {
    const written = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["test"]] }],
    });

    // Convert Uint8Array to ArrayBuffer
    const buffer = written.buffer.slice(
      written.byteOffset,
      written.byteOffset + written.byteLength,
    ) as ArrayBuffer;

    const workbook = await readOds(buffer);
    expect(workbook.sheets[0].rows[0]).toEqual(["test"]);
  });

  // ── Reader-specific spec compliance tests ─────────────────────────

  it("round-trips boolean cells preserving true/false values", async () => {
    const written = await writeOds({
      sheets: [{ name: "Bools", rows: [[true, false, true, false]] }],
    });

    const workbook = await readOds(written);
    const row = workbook.sheets[0].rows[0];
    expect(row).toEqual([true, false, true, false]);

    // Verify types
    expect(typeof row[0]).toBe("boolean");
    expect(typeof row[1]).toBe("boolean");
  });

  it("round-trips date cells as Date objects", async () => {
    const d1 = new Date(2026, 0, 15, 10, 0, 0);
    const d2 = new Date(2026, 5, 30, 23, 59, 59);

    const written = await writeOds({
      sheets: [{ name: "Dates", rows: [[d1, d2]] }],
    });

    const workbook = await readOds(written);
    const row = workbook.sheets[0].rows[0];
    expect(row[0]).toBeInstanceOf(Date);
    expect(row[1]).toBeInstanceOf(Date);
    expect((row[0] as Date).getFullYear()).toBe(2026);
    expect((row[1] as Date).getMonth()).toBe(5); // June = 5
  });

  it("round-trips all cell types in one row", async () => {
    const date = new Date(2026, 2, 25);
    const written = await writeOds({
      sheets: [
        {
          name: "AllTypes",
          rows: [["hello", 42, 3.14, true, false, date, null, "end"]],
        },
      ],
    });

    const workbook = await readOds(written);
    const row = workbook.sheets[0].rows[0];
    expect(row[0]).toBe("hello");
    expect(row[1]).toBe(42);
    expect(row[2]).toBe(3.14);
    expect(row[3]).toBe(true);
    expect(row[4]).toBe(false);
    expect(row[5]).toBeInstanceOf(Date);
    expect(row[6]).toBe(null);
    expect(row[7]).toBe("end");
  });

  it("handles rows with trailing nulls correctly (nulls trimmed)", async () => {
    const written = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["A", "B", null, null, null]] }],
    });

    const workbook = await readOds(written);
    // Trailing nulls should be trimmed
    expect(workbook.sheets[0].rows[0]).toEqual(["A", "B"]);
  });

  it("handles large sheet with sparse data efficiently", async () => {
    // Create a sheet with empty cells between data
    const rows: CellValue[][] = [
      ["A1", null, null, null, null, null, null, null, null, "J1"],
      [null, null, null, null, null, null, null, null, null, null], // fully empty row (should be trimmed)
      ["A3"],
    ];

    const written = await writeOds({
      sheets: [{ name: "Sparse", rows }],
    });

    const workbook = await readOds(written);
    expect(workbook.sheets[0].rows.length).toBe(2); // middle empty row trimmed from end? no, row 3 exists
    // Actually: row 0 has data, row 1 is empty (trimmed because all null), row 2 has data
    // The reader trims trailing null cells and skips fully-empty rows...
    // But with the new reader, fully empty rows between data rows are still skipped
    // Row 0: ["A1", null x 8, "J1"] -> ["A1", null, null, null, null, null, null, null, null, "J1"]
    // Row 1: all null -> skipped
    // Row 2: ["A3"] -> ["A3"]
    expect(workbook.sheets[0].rows[0][0]).toBe("A1");
    expect(workbook.sheets[0].rows[0][9]).toBe("J1");
  });

  it("reads cells with number-columns-repeated correctly", async () => {
    // Manually test that the reader handles repeated cells
    // by writing a file with repeated empty cells and reading it back
    const written = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["X", null, null, null, "Y"]] }],
    });

    const workbook = await readOds(written);
    const row = workbook.sheets[0].rows[0];
    expect(row[0]).toBe("X");
    expect(row[1]).toBe(null);
    expect(row[2]).toBe(null);
    expect(row[3]).toBe(null);
    expect(row[4]).toBe("Y");
    expect(row.length).toBe(5);
  });

  it("round-trips workbook with properties", async () => {
    const created = new Date("2026-01-01T00:00:00Z");
    const modified = new Date("2026-03-25T12:00:00Z");

    const written = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["data"]] }],
      properties: {
        title: "My Workbook",
        subject: "Test Data",
        creator: "Test Suite",
        description: "A test workbook",
        keywords: "test,ods,hucre",
        created,
        modified,
      },
    });

    const workbook = await readOds(written);
    expect(workbook.properties).toBeDefined();
    expect(workbook.properties!.title).toBe("My Workbook");
    expect(workbook.properties!.subject).toBe("Test Data");
    expect(workbook.properties!.creator).toBe("Test Suite");
    expect(workbook.properties!.description).toBe("A test workbook");
    expect(workbook.properties!.keywords).toBe("test,ods,hucre");
    expect(workbook.properties!.created).toBeInstanceOf(Date);
    expect(workbook.properties!.modified).toBeInstanceOf(Date);
  });

  it("round-trips sheet with only null cells (results in empty sheet)", async () => {
    const written = await writeOds({
      sheets: [{ name: "Nulls", rows: [[null, null, null]] }],
    });

    const workbook = await readOds(written);
    // Row of all nulls gets trimmed to empty, then the empty row is dropped
    expect(workbook.sheets[0].rows).toEqual([]);
  });

  it("round-trips negative numbers and zero", async () => {
    const written = await writeOds({
      sheets: [{ name: "Sheet1", rows: [[-1, 0, -99.99, 0.001]] }],
    });

    const workbook = await readOds(written);
    expect(workbook.sheets[0].rows[0]).toEqual([-1, 0, -99.99, 0.001]);
  });

  it("round-trips empty string cells", async () => {
    const written = await writeOds({
      sheets: [{ name: "Sheet1", rows: [["", "text", ""]] }],
    });

    const workbook = await readOds(written);
    const row = workbook.sheets[0].rows[0];
    expect(row[0]).toBe("");
    expect(row[1]).toBe("text");
    // Trailing empty string is not null, so it should be preserved
    expect(row[2]).toBe("");
  });

  it("round-trips special XML characters in strings", async () => {
    const written = await writeOds({
      sheets: [{ name: "Sheet1", rows: [['<script>alert("xss")</script>', "A & B", "1 > 0 < 2"]] }],
    });

    const workbook = await readOds(written);
    const row = workbook.sheets[0].rows[0];
    expect(row[0]).toBe('<script>alert("xss")</script>');
    expect(row[1]).toBe("A & B");
    expect(row[2]).toBe("1 > 0 < 2");
  });
});
